var util = require('util'), fs = require('fs'),
    EventEmitter = require('events').EventEmitter,
    Work = require('../../deps/work');

var Parser = module.exports = function(path, callback) {
  var self = this;
  this.fd = undefined;
  /*
    Header is always 512 bytes long, is always located at the beginning of the
    file, and occurs only once. It has the format (# of bytes in parens):

    magic bytes          (8) - always: [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1,
                                        0x1A, 0xE1]
    class id            (16)
    minor ver            (2) - minor version of format
    major ver            (2) - major version of format
    byte order           (2) - always [0xFF, 0xFE], indicating little endian
    sector size          (2) - size of sectors in power-of-two, usually 9,
                               so 2^9=512 byte sectors
    mini-sector size     (2) - size of mini-sectors in power-of-two, usually 6,
                               so 2^6=64 byte mini-sectors
    reserved            (10) - always zeroes
    SECTs in FAT         (4) - number of SECTs in the FAT chain
    SECT dir. start      (4) - first SECT in the directory chain
    transaction sig      (4) - signature used for transactioning (not supported)
                               -- must be zeroes
    mini-stream max size (4) - max size for mini-streams (less than but not
                               equal to), usually 4096 bytes
    mini-FAT chain start (4) - first SECT in mini-FAT chain
    SECTs in mini-FAT    (4) - number of SECTs in the mini-FAT chain
    DIF chain start      (4) - first SECT in DIF chain
    SECTs in DIF         (4) - number of SECTs in the DIF chain
    SECTs FAT sectors  (436) - SECTs of the first 109 FAT sectors

    SECT is simply a 4 byte unsigned long used to describe the location of a
    sector within a virtual stream (in most cases this virtual stream is the
    file itself).
  */
  this.header = undefined;
  this.FAT = undefined;
  this.DIF = undefined;
  this.miniFAT = undefined;
  this.miniStream = undefined;
  /*
    Directory entries have the format (# of bytes in parens):

    name                (64) - zero-padded unicode name
    name length          (2) - length of name in characters, not bytes
    type                 (1) - object type, one of STGTY_* values
    RB flags             (1) - 0 for red, 1 for black -- in context of RB trees
    left sibling         (4) - index of left sibling
    right sibling        (4) - index of right sibling
    child                (4) - index of root of children (STGTY_STORAGE)
    class id            (16) - (STGTY_STORAGE)
    user flags           (4) - (STGTY_STORAGE)
    create timestamp     (8) - (STGTY_STORAGE)
    modify timestamp     (8) - (STGTY_STORAGE)
    stream start         (4) - SECT of stream start (STGTY_STREAM)
    stream size          (8) - size of stream (STGTY_STREAM)
                               Note: for version 3 compound files with 512b
                               sectors, only the lower 32bits are to be used.
                               All 64 bits are only used for version 4 compound
                               files with 4096b sectors.

    Note: The root directory entry acts as both STGTY_STORAGE AND STGTY_STREAM.
  */
  this.dir = undefined;

  var work = new Work([
    function() { self._parseHeader(work.next.bind(work)); },
    function() { self._parseFAT(work.next.bind(work)); },
    function() {
      if (self.header.sectDIF !== ENDOFCHAIN)
        self._parseDIF(work.next.bind(work));
      else
        work.next();
    },
    function() {
      if (self.header.sectDir !== ENDOFCHAIN)
        self._parseDir(work.next.bind(work));
      else
        work.next();
    },
    function() {
      if (self.header.sectMiniFAT !== ENDOFCHAIN)
        self._parseMiniFAT(work.next.bind(work));
      else
        work.next();
    }
  ], callback);

  fs.open(path, 'r', function(err, fd) {
    if (err)
      return callback(err);
    self.fd = fd;
    work.go();
  });
};

Parser.prototype.close = function(cb) {
  fs.close(this.fd, cb);
};

Parser.prototype.findStream = function(name) {
  var ret = -1;
  name = name.toUpperCase();
  for (var i=0,len=this.dir.length; i<len; ++i) {
    if (this.dir[i].name.toUpperCase() === name) {
      ret = i;
      break;
    }
  }
  return ret;
};

Parser.prototype.getStream = function(id, cb) {
  var self = this, bytes, buf, sect, start, isFAT,
      streamSize, curSize = 0, stream;
  if (!this.dir[id] || this.dir[id].type !== STGTY_STREAM)
    return cb(new Error('There is no stream with that ID'));
  stream = new EventEmitter();
  isFAT = (this.dir[id].size >= this.header.maxMiniStreamSize);
  bytes = (isFAT ? this.header.sectorSize : this.header.miniSectorSize);
  buf = new Buffer(bytes);
  streamSize = this.dir[id].size;
  sect = this.dir[id].sect;
  var work = new Work(function() { stream.emit('end'); });
  while (sect !== ENDOFCHAIN) {
    curSize += bytes;
    if (isFAT)
      start = 512 + (sect * bytes);
    else
      start = 512 + (this.dir[0].sect * this.header.sectorSize + sect * bytes);
    work.push((function(pos) {
      return function() {
        fs.read(self.fd, buf, 0, bytes, pos, function(err, bytesRead) {
          if (err)
            return stream.emit('error', err);
          if (curSize > streamSize)
            stream.emit('data', buf.slice(0, (curSize-streamSize)));
          else
            stream.emit('data', buf);
          work.next();
        });
      }
    })(start));
    sect = (isFAT ? this.FAT[sect] : this.miniFAT[sect]);
  }
  cb(undefined, stream);
  work.go();
};

Parser.prototype._parseHeader = function(cb) {
  var buf = new Buffer(512), self = this;
  fs.read(this.fd, buf, 0, 512, 0, function(err, bytesRead) {
    if (err)
      return cb(err);
    else if (bytesRead !== 512)
      return cb(new Error('Invalid file format'));
    if (buf[0] !== 0xD0 || buf[1] !== 0xCF ||
        buf[2] !== 0x11 || buf[3] !== 0xE0 ||
        buf[4] !== 0xA1 || buf[5] !== 0xB1 ||
        buf[6] !== 0x1A || buf[7] !== 0xE1) {
      return cb(new Error('Invalid file format'));
    }
    self.header = {
      // skip magic bytes
      classId: buf.slice(8, 24).toArray(),
      version: {
        minor: buf.readUInt16(24, ENDIAN),
        major: buf.readUInt16(26, ENDIAN),
      },
      // skip byte order
      sectorSize: Math.pow(2, buf.readUInt16(30, ENDIAN)),
      miniSectorSize: Math.pow(2, buf.readUInt16(32, ENDIAN)),
      // skip reserved
      nSectFAT: buf.readUInt32(44, ENDIAN),
      sectDir: buf.readUInt32(48, ENDIAN),
      // skip transactioning signature
      maxMiniStreamSize: buf.readUInt32(56, ENDIAN),
      sectMiniFAT: buf.readUInt32(60, ENDIAN),
      nSectMiniFAT: buf.readUInt32(64, ENDIAN),
      sectDIF: buf.readUInt32(68, ENDIAN),
      nSectDIF: buf.readUInt32(72, ENDIAN),
      FATsects: new Array()
    };
    for (var i=76,j=-1,sect; i<436; i+=4) {
      sect = buf.readUInt32(i, ENDIAN);
      if (sect === ENDOFCHAIN || sect === FREESECT)
        break;
      self.header.FATsects[++j] = sect;
    }
    cb();
  });
};

Parser.prototype._parseFAT = function(cb) {
  var self = this, bytes = this.header.sectorSize, buf = new Buffer(bytes);
  var work = new Work(cb);
  for (var i=0,len=this.header.FATsects.length; i<len; ++i) {
    work.push((function(pos) {
      return function() {
        fs.read(self.fd, buf, 0, bytes, pos, function(err, bytesRead) {
          if (err)
            return cb(err);
          if (!self.FAT)
            self.FAT = new Array();
          for (var j=0,len=bytesRead; j<len; j+=4)
            self.FAT.push(buf.readUInt32(j, ENDIAN));
          work.next();
        });
      }
    })(512 + this.header.FATsects[i] * bytes));
  }
  work.go();
};

Parser.prototype._parseDir = function(cb) {
  var self = this, bytes = this.header.sectorSize, buf = new Buffer(bytes),
      sect = this.header.sectDir, entry, nEntries = bytes / 128;
  var work = new Work(cb);
  while (sect !== ENDOFCHAIN) {
    work.push((function(pos) {
      return function() {
        fs.read(self.fd, buf, 0, bytes, pos, function(err, bytesRead) {
          if (err)
            return cb(err);
          if (!self.dir)
            self.dir = new Array();
          for (var i=0,o; i<nEntries; ++i) {
            o = i*128;
            if (buf[o+66] === 0)
              break;
            entry = {
              name: buf.toString('utf8', o, o+63)
                       .substring(0, buf.readUInt16(o+64, ENDIAN) - 2)
                       .replace(/[\x00-\x1F]/g, ''),
              type: buf[o+66],
              left: buf.readUInt32(o+68, ENDIAN),
              right: buf.readUInt32(o+72, ENDIAN)
            };
            self.dir.push(entry);
            if (entry.type === STGTY_STORAGE || entry.type === STGTY_ROOT) {
              entry.children = undefined;
              entry.child = buf.readUInt32(o+76, ENDIAN);
              entry.classId = makeClsId(buf.slice(o+80, o+96));
              entry.userFlags = buf.readUInt32(o+96, ENDIAN);
              entry.createTS = buf.readUInt32(o+100, ENDIAN);
              entry.modifyTS = buf.readUInt32(o+108, ENDIAN);
            }
            if (entry.type === STGTY_STREAM || entry.type === STGTY_ROOT) {
              entry.sect = buf.readUInt32(o+116, ENDIAN);
              if (self.header.version.major === 3 && bytes === 512)
                entry.size = buf.readUInt32(o+120, ENDIAN);
              else if (self.header.version.major === 4 && bytes === 4096)
                entry.size = buf.readDouble(o+120, ENDIAN);
              else
                throw new Error('Unsure of mini stream size. Version === '
                                 + self.header.version.major + ', Sector size: '
                                 + bytes);
              if (buf[o] === 5) {
                // this stream has a property set
                (function(ixEntry, size) {
                  // HACK
                  work._tasks.push(function() {
                    self.getStream(ixEntry, function(err, stream) {
                      if (err)
                        return work.next();
                      var bufProps = new Buffer(size), offset = 0;
                      stream.on('data', function(data) {
                        data.copy(bufProps, offset);
                        offset += data.length;
                      });
                      stream.on('error', function(err) {
                        work.next();
                      });
                      stream.on('end', function() {
                        /*
                          First 28 bytes of bufProps is a "Property Set Header"
                          with the structure of PROPERTYSETHEADER:

                            byte order      (2) - Always 0xFFFE (little endian)
                            format version  (2) - 0 or 1. Version 1 is equiv.
                                                  to version 0 except:
                                                    * property id 0 property
                                                      names can be case-
                                                      sensitive, depending on
                                                      the value of the reserved
                                                      Behavior property in
                                                      property id 0x80000003
                                                    * property id 0 property
                                                      names can have a count
                                                      greater than 256 (bytes
                                                      or chars depending on if
                                                      the property set's
                                                      codepage is set to Unicode
                                                      or not)
                                                    * more property types have
                                                      been added
                            OS version      (4) - System version
                            class ID       (16) - Application CLSID
                            prop. set count (4) - Should be 1 or 2 (sections
                                                  start with
                                                  PROPERTYSECTIONHEADER)

                          Next 20 bytes has the structure of FORMATIDOFFSET:

                            format ID      (16) - Unique ID representing the
                                                  first property set
                            offset start    (4) - offset for start of the first
                                                  property set
                            format Id      (16) - Unique ID representing the
                                                  second property set (if avail)
                            offset start    (4) - offset for start of the second
                                                  property set (if avail)

                          Next 8 bytes has the structure of
                          PROPERTYSECTIONHEADER:

                            total byte size (4) - total size of the entire set
                                                  including this byte count.
                                                  must be at least 262,144b and
                                                  should be 2,097,152b.
                            property count  (4) - total number of properties in
                                                  this set

                          Next is an array of PROPERTYIDOFFSET structures:

                            property id     (4) - id unique to this particular
                                                  set
                            property offset (4) - offset of property info
                                                  relative to start of set

                          Next is an array of SERIALIZEDPROPERTYVALUE
                          structures:

                            property type   (4) - VT_* constant
                            property data   (?) - size and contents of this
                                                  field depends on the property
                                                  type
                        */
                        var props = self.dir[ixEntry].properties = new Object(),
                            start = bufProps.readUInt32(44, ENDIAN),
                            numProps = bufProps.readUInt32(start+4, ENDIAN),
                            loc, prop, c, type, id;
                        props.fmtVer = bufProps.readUInt16(2, ENDIAN);
                        props.fmtId = makeClsId(bufProps.slice(28, 44));
                        props.items = new Array();
                        for (var i=0; i<numProps; ++i) {
                          id = bufProps.readUInt32(start+i*8+8, ENDIAN);
                          loc = bufProps.readUInt32(start+i*8+12, ENDIAN);
                          loc += start;
                          type = bufProps.readUInt32(loc, ENDIAN);
                          props.items.push(prop = {
                            id: id,
                            type: type
                          });
                          if (type === VT_I1) {
                            // 8-bit signed int
                            prop.value = bufProps.readInt8(loc+4, ENDIAN);
                          } else if (type === VT_UI1) {
                            // 8-bit unsigned int
                            prop.value = bufProps.readUInt8(loc+4, ENDIAN);
                          } else if (type === VT_I2) {
                            // 16-bit signed int
                            prop.value = bufProps.readInt16(loc+4, ENDIAN);
                            if (prop.value >= 32768)
                              prop.value -= 65536;
                          } else if (type === VT_UI2) {
                            // 16-bit unsigned int
                            prop.value = bufProps.readUInt16(loc+4, ENDIAN);
                          } else if (type === VT_I4 || type === VT_ERROR ||
                                     type === VT_INT) {
                            // 32-bit signed int
                            prop.value = bufProps.readInt32(loc+4, ENDIAN);
                          } else if (type === VT_UI4 || type === VT_UINT) {
                            // 32-bit unsigned int
                            prop.value = bufProps.readUInt32(loc+4, ENDIAN);
                          } else if (type === VT_R4) {
                            // 32-bit float
                            prop.value = bufProps.readFloat(loc+4, ENDIAN);
                          } else if (type === VT_R8) {
                            // 64-bit double
                            prop.value = bufProps.readDouble(loc+4, ENDIAN);
                          } else if (type === VT_BSTR) {
                            // binary string terminated with double null bytes.
                            // encoding depends on property set's codepage
                            // property
                            c = bufProps.readUInt32(loc+4, ENDIAN);
                            prop.value = new Buffer(c);
                            bufProps.copy(prop.value, 0, loc+8, loc+8+c-1);
                          } else if (type === VT_LPSTR) {
                            // 8-bit ANSI string
                            c = bufProps.readUInt32(loc+4, ENDIAN);
                            var s = loc+8;
                            prop.value = bufProps.toString('utf8', s, s+c-1);
                          } else if (type === VT_BLOB) {
                            // binary blob
                            c = bufProps.readUInt32(loc+4, ENDIAN);
                            prop.value = new Buffer(c);
                            bufProps.copy(prop.value, 0, loc+8, loc+8+c);
                          } else if (type === VT_LPWSTR) {
                            // utf-16 string
                            c = bufProps.readUInt32(loc+4, ENDIAN);
                            var s = loc+8;
                            prop.value = bufProps.toString('ucs2', s, s+c*2);
                          } else if (type === VT_DATE) {
                            // 64-bit double (same as VT_R8) of the number of
                            // days since 12/31/1899
                            var val = bufProps.readDouble(loc+4, ENDIAN),
                                unixDays = Date.now() / 86400;
                            // convert to UNIX timestamp
                            prop.value = (val - (val - unixDays)) * 86400;
                          } else if (type === VT_BOOL) {
                            prop.value = (bufProps[loc+4] === 0 ? false : true);
                          } else if (type === VT_FILETIME) {
                            // 64-bit FILETIME structure
                            // Represents the number of 100-nanosecond intervals
                            // since January 1, 1601 (UTC)
                            var high = bufProps.readUInt32(loc+8, 'little'),
                                low = bufProps.readUInt32(loc+4, 'little');
                            high = lshift(high, 32);
                            if (id === PID_EDITTIME)
                              prop.value = (high + low) / 10000000; // seconds
                            else
                              prop.value = (high + low - 116444736000000000)
                                           / 10000000;
                          } else if (type === VT_CLSID) {
                            prop.value = makeClsId(bufProps.slice(loc+4, loc+20));
                          } else if (type === VT_NULL)
                            prop.value = null;
                        }
                        work.next();
                      });
                    });
                  });
                })(self.dir.length-1, entry.size);
              }
            }
          }
          work.next();
        });
      }
    })(512 + sect * bytes));
    sect = this.FAT[sect];
  }
  work.push(function() {
    for (var i=0,cur,ids,len=self.dir.length; i<len; ++i) {
      if (typeof self.dir[i].child !== 'undefined'
          && self.dir[i].child !== -1) {
        self.dir[i].children = new Array();
        ids = [self.dir[i].child];
        while (ids.length) {
          cur = ids.pop();
          if (cur !== FREESECT) {
            if (self.dir[cur].left !== FREESECT)
              ids.push(self.dir[cur].left);
            if (self.dir[cur].right !== FREESECT)
              ids.push(self.dir[cur].right);
            self.dir[i].children.push(self.dir[cur]);
          }
        }
      }
      delete self.dir[i].child;
      delete self.dir[i].left;
      delete self.dir[i].right;
      delete self.dir[i].type;
    }
    self.dir = self.dir[0];
    work.next();
  });
  work.go();
};

Parser.prototype._parseMiniFAT = function(cb) {
  var self = this, bytes = this.header.sectorSize, buf = new Buffer(bytes),
      sect = this.header.sectMiniFAT;
  var work = new Work(cb);
  while (sect !== ENDOFCHAIN) {
    work.push((function(pos) {
      return function() {
        fs.read(self.fd, buf, 0, bytes, pos, function(err, bytesRead) {
          if (err)
            return cb(err);
          if (!self.miniFAT)
            self.miniFAT = new Array();
          for (var j=0,len=bytesRead; j<len; j+=4)
            self.miniFAT.push(buf.readUInt32(j, ENDIAN));
          work.next();
        });
      }
    })(512 + sect * bytes));
    sect = this.FAT[sect];
  }
  work.go();
};

Parser.prototype._parseDIF = function(cb) {
  var self = this, bytes = this.header.sectorSize, buf = new Buffer(bytes),
      sect = this.header.sectDIF, pos = 512 + sect * bytes;
  var work = new Work(cb);
  work.push(function() {
    var thisFn = this;
    fs.read(self.fd, buf, 0, bytes, pos, function(err, bytesRead) {
      if (err)
        return cb(err);
      var lastByte = bytesRead - 4;
      for (var j=0,len=bytesRead; j<len; j+=4) {
        sect = buf.readUInt32(j, ENDIAN);
        if (j === lastByte) {
          if (sect !== ENDOFCHAIN) {
            pos = 512 + sect * bytes;
            work._tasks.push(thisFn);
          }
        } else
          self.FAT.push(sect);
      }
      work.next();
    });
  });
  work.go();
};

/* Constants */

var ENDIAN = 'little';

// CLSID for root directory entry
var CLSID = {
  EXCEL: [
    // 97+
    [0x00,0x02,0x08,0x12,0x00,0x00,0x00,0x00,
     0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46],
    // 95
    [0x00,0x02,0x08,0x10,0x00,0x00,0x00,0x00,
     0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46]
  ],
  WORD: [
    // 97+
    [0x00,0x02,0x09,0x06,0x00,0x00,0x00,0x00,
     0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46],
    // 95
    [0x00,0x02,0x09,0x00,0x00,0x00,0x00,0x00,
     0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46]
  ],
  POWERPOINT: [
    // 97+
    [0x64,0x81,0x8D,0x10,0x4F,0x9B,0x11,0xCF,
     0x86,0xEA,0x00,0xAA,0x00,0xB9,0x29,0xE8]
  ],
  ACCESS: [
    // 97
    [0x8C,0xC4,0x99,0x40,0x31,0x46,0x11,0xCF,
     0x97,0xA1,0x00,0xAA,0x00,0x42,0x4A,0x9F],
    // 2000/2002
    [0x73,0xA4,0xC9,0xC1,0xD6,0x8D,0x11,0xD0,
     0x98,0xBF,0x00,0xA0,0xC9,0x0D,0xC8,0xD9]
  ]
}

var FORMATID = {
  SUMMARY: [0xF2, 0x9F, 0x85, 0xE0, 0x4F, 0xF9, 0x10, 0x68,
            0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9],
  DOCSUMMARY: [0xD5, 0xCD, 0xD5, 0x02, 0x2E, 0x9C, 0x10, 0x1B,
               0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE],
  USERDEFPROPS: [0xD5, 0xCD, 0xD5, 0x05, 0x2E, 0x9C, 0x10, 0x1B,
               0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE],
  GLOBALINFO: [0x56, 0x61, 0x6F, 0x00, 0xC1, 0x54, 0x11, 0xCE,
               0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B],
  IMAGECONTENTS: [0x56, 0x61, 0x64, 0x00, 0xC1, 0x54, 0x11, 0xCE,
                  0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B],
  IMAGEINFO: [0x56, 0x61, 0x65, 0x00, 0xC1, 0x54, 0x11, 0xCE,
               0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B]
};

// Special FAT entry values
var DIFSECT    = 0xFFFFFFFC,
    FATSECT    = 0xFFFFFFFD,
    ENDOFCHAIN = 0xFFFFFFFE,
    FREESECT   = 0xFFFFFFFF;

// Directory sector object types
var STGTY_INVALID   = 0,
    STGTY_STORAGE   = 1,
    STGTY_STREAM    = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY  = 4,
    STGTY_ROOT      = 5;

// Stream property types
var VT_EMPTY           = 0,
    VT_NULL            = 1,
    VT_I2              = 2,
    VT_I4              = 3,
    VT_R4              = 4,
    VT_R8              = 5,
    VT_CY              = 6,
    VT_DATE            = 7,
    VT_BSTR            = 8,
    VT_DISPATCH        = 9,
    VT_ERROR           = 10,
    VT_BOOL            = 11,
    VT_VARIANT         = 12,
    VT_UNKNOWN         = 13,
    VT_DECIMAL         = 14,
    VT_I1              = 16,
    VT_UI1             = 17,
    VT_UI2             = 18,
    VT_UI4             = 19,
    VT_I8              = 20,
    VT_UI8             = 21,
    VT_INT             = 22,
    VT_UINT            = 23,
    VT_VOID            = 24,
    VT_HRESULT         = 25,
    VT_PTR             = 26,
    VT_SAFEARRAY       = 27,
    VT_CARRAY          = 28,
    VT_USERDEFINED     = 29,
    VT_LPSTR           = 30,
    VT_LPWSTR          = 31,
    VT_FILETIME        = 64,
    VT_BLOB            = 65,
    VT_STREAM          = 66,
    VT_STORAGE         = 67,
    VT_STREAMED_OBJECT = 68,
    VT_STORED_OBJECT   = 69,
    VT_BLOB_OBJECT     = 70,
    VT_CF              = 71,
    VT_CLSID           = 72,
    VT_VECTOR          = 4096;

// Well-known property ids
var PID_APPNAME           = 18, // SummaryInformation
    PID_AUTHOR            = 4,  // SummaryInformation
    PID_BEHAVIOR          = 0x80000003, // format version 1 only. Value is 0
                                        // (default) for case-insensitive
                                        // property names, 1 for case-sensitive
    PID_BYTECOUNT         = 4,  // DocumentSummaryInformation
    PID_CATEGORY          = 2,  // DocumentSummaryInformation
    PID_CCHWITHSPACES     = 17, // DocumentSummaryInformation
    PID_CHARCOUNT         = 16, // SummaryInformation
    PID_CODEPAGE          = 1,
    PID_COMMENTS          = 6,  // SummaryInformation
    PID_COMPANY           = 15, // DocumentSummaryInformation
    PID_CONTENTSTATUS     = 27, // DocumentSummaryInformation
    PID_CONTENTTYPE       = 26, // DocumentSummaryInformation
    PID_CREATE_DTM        = 12, // SummaryInformation
    PID_DICTIONARY        = 0,
    PID_DIGSIG            = 24, // DocumentSummaryInformation
    PID_DOCPARTS          = 13, // DocumentSummaryInformation
    PID_DOCVERSION        = 29, // DocumentSummaryInformation
    PID_EDITTIME          = 10, // SummaryInformation
    PID_HEADINGPAIR       = 12, // DocumentSummaryInformation
    PID_HIDDENCOUNT       = 9,  // DocumentSummaryInformation
    PID_HLINKS            = 21, // DocumentSummaryInformation
    PID_HYPERLINKSCHANGED = 22, // DocumentSummaryInformation
    PID_KEYWORDS          = 5,  // SummaryInformation
    PID_LANGUAGE          = 28, // DocumentSummaryInformation
    PID_LASTAUTHOR        = 8,  // SummaryInformation
    PID_LASTPRINTED       = 11, // SummaryInformation
    PID_LASTSAVE_DTM      = 13, // SummaryInformation
    PID_LINECOUNT         = 5,  // DocumentSummaryInformation
    PID_LINKBASE          = 20, // DocumentSummaryInformation
    PID_LINKSDIRTY        = 16, // DocumentSummaryInformation
    PID_LOCALE            = 0x80000000,
    PID_MANAGER           = 14, // DocumentSummaryInformation
    PID_MAX               = 16,
    PID_MMCLIPCOUNT       = 10, // DocumentSummaryInformation
    PID_NOTECOUNT         = 8,  // DocumentSummaryInformation
    PID_PAGECOUNT         = 14, // SummaryInformation
    PID_PARCOUNT          = 6,  // DocumentSummaryInformation
    PID_PRESFORMAT        = 3,  // DocumentSummaryInformation
    PID_REVNUMBER         = 9,  // SummaryInformation
    PID_SCALE             = 11, // DocumentSummaryInformation
    PID_SECURITY          = 19, // SummaryInformation. Bit field values:
                                //  0 - no security
                                //  1 - Password protected
                                //  2 - read-only recommended
                                //  4 - read-only enforced
                                //  8 - locked for annotations
    PID_SHAREDDOC         = 19, // DocumentSummaryInformation
    PID_SLIDECOUNT        = 7,  // DocumentSummaryInformation
    PID_SUBJECT           = 3,  // SummaryInformation
    PID_TEMPLATE          = 7,  // SummaryInformation
    PID_THUMBNAIL         = 17, // SummaryInformation
    PID_TITLE             = 2,  // SummaryInformation
    PID_VERSION           = 23, // DocumentSummaryInformation
    PID_WORDCOUNT         = 15; // SummaryInformation

// Codepages for stream property id 1
var CP_037                     = 37,
    CP_EUC_JP                  = 51932,
    CP_EUC_KR                  = 51949,
    CP_GB18030                 = 54936,
    CP_GB2312                  = 52936,
    CP_GBK                     = 936,
    CP_ISO_2022_JP1            = 50220,
    CP_ISO_2022_JP2            = 50221,
    CP_ISO_2022_JP3            = 50222,
    CP_ISO_2022_KR             = 50225,
    CP_ISO_8859_1              = 28591,
    CP_ISO_8859_2              = 28592,
    CP_ISO_8859_3              = 28593,
    CP_ISO_8859_4              = 28594,
    CP_ISO_8859_5              = 28595,
    CP_ISO_8859_6              = 28596,
    CP_ISO_8859_7              = 28597,
    CP_ISO_8859_8              = 28598,
    CP_ISO_8859_9              = 28599,
    CP_JOHAB                   = 1361,
    CP_KOI8_R                  = 20866,
    CP_MAC_ARABIC              = 10004,
    CP_MAC_CENTRAL_EUROPE      = 10029,
    CP_MAC_CHINESE_SIMPLE      = 10008,
    CP_MAC_CHINESE_TRADITIONAL = 10002,
    CP_MAC_CROATIAN            = 10082,
    CP_MAC_CYRILLIC            = 10007,
    CP_MAC_GREEK               = 10006,
    CP_MAC_HEBREW              = 10005,
    CP_MAC_ICELAND             = 10079,
    CP_MAC_JAPAN               = 10001,
    CP_MAC_KOREAN              = 10003,
    CP_MAC_ROMAN               = 10000,
    CP_MAC_ROMANIA             = 10010,
    CP_MAC_THAI                = 10021,
    CP_MAC_TURKISH             = 10081,
    CP_MAC_UKRAINE             = 10017,
    CP_MS949                   = 949,
    CP_SJIS                    = 932,
    CP_UNICODE                 = 1200,
    CP_US_ACSII                = 20127,
    CP_US_ASCII2               = 65000,
    CP_UTF16                   = 1200,
    CP_UTF16_BE                = 1201,
    CP_UTF8                    = 65001,
    CP_WINDOWS_1250            = 1250,
    CP_WINDOWS_1251            = 1251,
    CP_WINDOWS_1252            = 1252,
    CP_WINDOWS_1253            = 1253,
    CP_WINDOWS_1254            = 1254,
    CP_WINDOWS_1255            = 1255,
    CP_WINDOWS_1256            = 1256,
    CP_WINDOWS_1257            = 1257,
    CP_WINDOWS_1258            = 1258;

/* Utility functions */

function lshift(num, bits) {
  return num * Math.pow(2, bits);
}

function makeClsId(buf) {
  var clsid = new Array(16);
  clsid[0] = buf[3];
  clsid[1] = buf[2];
  clsid[2] = buf[1];
  clsid[3] = buf[0];

  clsid[4] = buf[5];
  clsid[5] = buf[4];

  clsid[6] = buf[7];
  clsid[7] = buf[6];

  for (var i=8; i<16; ++i)
    clsid[i] = buf[i];
  return clsid;
}

Buffer.prototype.toArray = function() {
  var len = this.length, a = new Array(len);
  for (var i=0; i<len; ++i)
    a[i] = this[i];
  return a;
};

if (typeof Buffer.prototype.readUInt32 === 'undefined') {
  Buffer.prototype.readUInt32 = function(offset, endian) {
    var val = 0;

    if (endian == 'big') {
      val = this[offset + 1] << 16;
      val |= this[offset + 2] << 8;
      val |= this[offset + 3];
      val = val + (this[offset] << 24 >>> 0);
    } else {
      val = this[offset + 2] << 16;
      val |= this[offset + 1] << 8;
      val |= this[offset];
      val = val + (this[offset + 3] << 24 >>> 0);
    }

    return val;
  };

  Buffer.prototype.readUInt16 = function(offset, endian) {
    var val = 0;

    if (endian == 'big') {
      val = this[offset] << 8;
      val |= this[offset + 1];
    } else {
      val = this[offset];
      val |= this[offset + 1] << 8;
    }

    return val;
  };


  Buffer.prototype.readUInt32 = function(offset, endian) {
    var val = 0;

    if (endian == 'big') {
      val = this[offset + 1] << 16;
      val |= this[offset + 2] << 8;
      val |= this[offset + 3];
      val = val + (this[offset] << 24 >>> 0);
    } else {
      val = this[offset + 2] << 16;
      val |= this[offset + 1] << 8;
      val |= this[offset];
      val = val + (this[offset + 3] << 24 >>> 0);
    }

    return val;
  };

  Buffer.prototype.readInt8 = function(offset, endian) {
    if (!(this[offset] & 0x80))
      return (this[offset]);

    return ((0xff - this[offset] + 1) * -1);
  };


  Buffer.prototype.readInt16 = function(offset, endian) {
    var val = this.readUInt16(offset, endian);
    if (!(val & 0x8000))
      return val;

    return (0xffff - val + 1) * -1;
  };

  Buffer.prototype.readInt32 = function(offset, endian) {
    var val = this.readUInt32(offset, endian);
    if (!(val & 0x80000000))
      return (val);

    return (0xffffffff - val + 1) * -1;
  };

  Buffer.prototype.readFloat = function(offset, endian) {
    return readIEEE754(this, offset, endian, 23, 4);
  };

  Buffer.prototype.readDouble = function(offset, endian) {
    return readIEEE754(this, offset, endian, 52, 8);
  };

  function readIEEE754(buffer, offset, endian, mLen, nBytes) {
    var e, m,
        bBE = (endian === 'big'),
        eLen = nBytes * 8 - mLen - 1,
        eMax = (1 << eLen) - 1,
        eBias = eMax >> 1,
        nBits = -7,
        i = bBE ? 0 : (nBytes - 1),
        d = bBE ? 1 : -1,
        s = buffer[offset + i];

    i += d;

    e = s & ((1 << (-nBits)) - 1);
    s >>= (-nBits);
    nBits += eLen;
    for (; nBits > 0; e = e * 256 + buffer[offset + i], i += d, nBits -= 8);

    m = e & ((1 << (-nBits)) - 1);
    e >>= (-nBits);
    nBits += mLen;
    for (; nBits > 0; m = m * 256 + buffer[offset + i], i += d, nBits -= 8);

    if (e === 0) {
      e = 1 - eBias;
    } else if (e === eMax) {
      return m ? NaN : ((s ? -1 : 1) * Infinity);
    } else {
      m = m + Math.pow(2, mLen);
      e = e - eBias;
    }
    return (s ? -1 : 1) * m * Math.pow(2, e - mLen);
  }
}

Array.prototype.equals = function(arr) {
  var ret = true;
  if (this.length === arr.length) {
    for (var i=0,len=this.length; i<len; ++i) {
      if (this[i] !== arr[i]) {
        ret = false;
        break;
      }
    }
  } else
    ret = false;
  return ret;
};