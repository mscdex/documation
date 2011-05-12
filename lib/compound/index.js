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
    stream size          (4) - size of stream (STGTY_STREAM)
    reserved             (2) - must be zero

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
    function() { self._parseDir(work.next.bind(work)); },
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
      start = 512 + ((this.dir[0].sect * this.header.sectorSize) + (sect * bytes));
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
        minor: buf.getIntLE(26, 2),
        major: buf.getIntLE(24, 2),
      },
      // skip byte order
      sectorSize: Math.pow(2, buf.getIntLE(30, 2)),
      miniSectorSize: Math.pow(2, buf.getIntLE(32, 2)),
      // skip reserved
      nSectFAT: buf.getIntLE(44, 4),
      sectDir: buf.getIntLE(48, 4),
      // skip transactioning signature
      maxMiniStreamSize: buf.getIntLE(56, 4),
      sectMiniFAT: buf.getIntLE(60, 4),
      nSectMiniFAT: buf.getIntLE(64, 4),
      sectDIF: buf.getIntLE(68, 4),
      nSectDIF: buf.getIntLE(72, 4),
      FATsects: new Array()
    };
    for (var i=76,j=-1,sect; i<436; i+=4) {
      sect = buf.getIntLE(i, 4);
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
            self.FAT.push(buf.getIntLE(j, 4));
          work.next();
        });
      }
    })(512 + (this.header.FATsects[i] * bytes)));
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
                       .substring(0, buf.getIntLE(o+64, 2)-2)
                       .replace(/[\x00-\x1F]/g, ''),
              type: buf[o+66],
              left: buf.getIntLE(o+68, 4),
              right: buf.getIntLE(o+72, 4)
            };
            self.dir.push(entry);
            if (entry.type === STGTY_STORAGE || entry.type === STGTY_ROOT) {
              entry.child = buf.getIntLE(o+76, 4);
              entry.classId = makeClsId(buf.slice(o+80, o+96));
              entry.userFlags = buf.getIntLE(o+96, 4);
              entry.createTS = buf.getIntLE(o+100, 8);
              entry.modifyTS = buf.getIntLE(o+108, 8);
            }
            if (entry.type === STGTY_STREAM || entry.type === STGTY_ROOT) {
              entry.sect = buf.getIntLE(o+116, 4);
              entry.size = buf.getIntLE(o+120, 4);
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
                                                      the property set's codepage
                                                      is set to Unicode or not)
                                                    * more property types have
                                                      been added
                            OS version      (4) - System version
                            class ID       (16) - Application CLSID
                            section count   (4) - Should be 1 (sections start
                                                  with PROPERTYSECTIONHEADER)
                          Next 20 bytes has the structure of FORMATIDOFFSET:
                            format ID      (16) - Unique ID representing this
                                                  property set
                            offset start    (4) - offset for start of actual
                                                  property set
                          Next 8 bytes has the structure of
                          PROPERTYSECTIONHEADER:
                            total byte size (4) - total size of the entire set
                                                  including this byte count
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
                            start = bufProps.getIntLE(44, 4),
                            numProps = bufProps.getIntLE(start+4, 4),
                            loc, prop, count, type, id;
                        props.fmtVer = bufProps.getIntLE(2, 2);
                        props.fmtId = makeClsId(bufProps.slice(28, 44));
                        props.items = new Array();
                        for (var i=0; i<numProps; ++i) {
                          id = bufProps.getIntLE(start+(i*8)+8, 4);
                          loc = start + bufProps.getIntLE(start+(i*8+12), 4);
                          type = bufProps.getIntLE(loc, 4);
                          props.items.push(prop = {
                            id: id,
                            type: type
                          });
                          if (type === VT_I1) {
                            // 8-bit signed int
                            prop.value = bufProps[loc+4];
                            if (prop.value >= 128)
                              prop.value -= 256;
                          } else if (type === VT_UI1) {
                            // 8-bit unsigned int
                            prop.value = bufProps[loc+4];
                          } else if (type === VT_I2) {
                            // 16-bit signed int
                            prop.value = bufProps.getIntLE(loc+4, 2);
                            if (prop.value >= 32768)
                              prop.value -= 65536;
                          } else if (type === VT_UI2) {
                            // 16-bit unsigned int
                            prop.value = bufProps.getIntLE(loc+4, 2).toUnsigned();
                          } else if (type === VT_I4 || type === VT_ERROR ||
                                     type === VT_INT) {
                            // 32-bit signed int
                            prop.value = bufProps.getIntLE(loc+4, 4);
                          } else if (type === VT_UI4 || type === VT_UINT) {
                            // 32-bit unsigned int
                            prop.value = bufProps.getIntLE(loc+4, 4).toUnsigned();
                          } else if (type === VT_R4) {
                            // 32-bit float
                            prop.value = bufProps.getDoubleLE(loc+4, 4);
                          } else if (type === VT_R8) {
                            // 64-bit double
                            prop.value = bufProps.getDoubleLE(loc+4, 8);
                          } else if (type === VT_BSTR) {
                            // binary string terminated with double null bytes
                            count = bufProps.getIntLE(loc+4, 4);
                            prop.value = new Buffer(count);
                            bufProps.copy(prop.value, 0, loc+8, loc+8+count-1);
                          } else if (type === VT_LPSTR) {
                            // 8-bit ANSI string
                            count = bufProps.getIntLE(loc+4, 4);
                            prop.value = bufProps.toString('utf8', loc+8, loc+8+count-1);
                          } else if (type === VT_BLOB) {
                            // binary blob
                            count = bufProps.getIntLE(loc+4, 4);
                            prop.value = new Buffer(count);
                            bufProps.copy(prop.value, 0, loc+8, loc+8+count);
                          } else if (type === VT_LPWSTR) {
                            // utf-16 string
                            count = bufProps.getIntLE(loc+4, 4);
                            prop.value = bufProps.toString('ucs2', loc+8, loc+8+count*2);
                          } else if (type === VT_DATE) {
                            // 64-bit double (same as VT_R8) of the number of
                            // days since 12/31/1899
                            var val = bufProps.getDoubleLE(loc+4, 8),
                                unixDays = Date.now() / 86400;
                            // convert to UNIX timestamp
                            prop.value = (val - (val - unixDays)) * 86400;
                          } else if (type === VT_BOOL) {
                            prop.value = (bufProps[loc+4] === 0 ? false : true);
                          } else if (type === VT_FILETIME) {
                            // 64-bit FILETIME structure
                            // Represents the number of 100-nanosecond intervals
                            // since January 1, 1601 (UTC)
                            var high = lshift(bufProps.readUInt32(loc+8, 'little'), 32),
                                low = bufProps.readUInt32(loc+4, 'little');
                            prop.value = (((high + low) - 116444736000000000)
                                          / 10000000);
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
    })(512 + (sect * bytes)));
    sect = this.FAT[sect];
  }
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
            self.miniFAT.push(buf.getIntLE(j, 4));
          work.next();
        });
      }
    })(512 + (sect * bytes)));
    sect = this.FAT[sect];
  }
  work.go();
};

Parser.prototype._parseDIF = function(cb) {
  // TODO
};

/* Constants */

/*var CLSID = {
  EXCEL97: [0x00,0x02,0x08,0x20,0x00,0x00,0x00,0x00,
            0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46],

  EXCEL95: [0x00,0x02,0x08,0x10,0x00,0x00,0x00,0x00,
            0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46],

  WORD97: [0x00,0x02,0x09,0x06,0x00,0x00,0x00,0x00,
           0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46],

  WORD95: [0x00,0x02,0x09,0x00,0x00,0x00,0x00,0x00,
           0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46]
}*/

var FORMATID = {
  SUMMARY: [0xF2, 0x9F, 0x85, 0xE0, 0x4F, 0xF9, 0x10, 0x68,
            0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9],
  DOCSUMMARY: [0xD5, 0xCD, 0xD5, 0x02, 0x2E, 0x9C, 0x10, 0x1B,
               0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE]
};

// Special FAT entry values
var DIFSECT    = -4,
    FATSECT    = -3,
    ENDOFCHAIN = -2,
    FREESECT   = -1;

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
var PID_APPNAME      = 18,
    PID_AUTHOR       = 4,
    PID_BYTECOUNT    = 4,
    PID_CATEGORY     = 2,
    PID_CHARCOUNT    = 16,
    PID_CODEPAGE     = 1,
    PID_COMMENTS     = 6,
    PID_COMPANY      = 15,
    PID_CREATE_DTM   = 12,
    PID_DICTIONARY   = 0,
    PID_DOCPARTS     = 13,
    PID_EDITTIME     = 10,
    PID_HEADINGPAIR  = 12,
    PID_HIDDENCOUNT  = 9,
    PID_KEYWORDS     = 5,
    PID_LASTAUTHOR   = 8,
    PID_LASTPRINTED  = 11,
    PID_LASTSAVE_DTM = 13,
    PID_LINECOUNT    = 5,
    PID_LINKSDIRTY   = 16,
    PID_MANAGER      = 14,
    PID_MAX          = 16,
    PID_MMCLIPCOUNT  = 10,
    PID_NOTECOUNT    = 8,
    PID_PAGECOUNT    = 14,
    PID_PARCOUNT     = 6,
    PID_PRESFORMAT   = 3,
    PID_REVNUMBER    = 9,
    PID_SCALE        = 11,
    PID_SECURITY     = 19,
    PID_SLIDECOUNT   = 7,
    PID_SUBJECT      = 3,
    PID_TEMPLATE     = 7,
    PID_THUMBNAIL    = 17,
    PID_TITLE        = 2,
    PID_WORDCOUNT    = 15,
    PID_LOCALE       = -0x80000000,
    PID_BEHAVIOR     = -0x80000003;

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

Number.prototype.toUnsigned = function() {
	return ((this >>> 1) * 2 + (this & 1));
};

Buffer.prototype.toArray = function() {
  var a = new Array(this.length);
  for (var i=0,len=this.length; i<len; ++i)
    a[i] = this[i];
  return a;
};

Buffer.prototype.getIntLE = function(index, bytes) {
  bytes || (bytes = 1);
  var sum = 0, shift = 0;
  for (var i=index,len=index+bytes; i<len; ++i) {
    sum += (this[i] << shift);
    shift += 8;
  }
  return sum;
};

Buffer.prototype.readUInt32 = function(offset, endian) {
  var val = 0;
  var buffer = this;

  if (endian == 'big') {
    val = buffer[offset + 1] << 16;
    val |= buffer[offset + 2] << 8;
    val |= buffer[offset + 3];
    val = val + (buffer[offset] << 24 >>> 0);
  } else {
    val = buffer[offset + 2] << 16;
    val |= buffer[offset + 1] << 8;
    val |= buffer[offset];
    val = val + (buffer[offset + 3] << 24 >>> 0);
  }

  return val;
};

Buffer.prototype.getDoubleLE = function(index, bytes) {
  bytes || (bytes = 8);
  var s, e, m, i, d, nBits, mLen, eLen, eBias, eMax;
  mLen = (bytes === 8 ? 52 : 23), eLen = bytes*8-mLen-1, eMax = (1<<eLen)-1,
  eBias = eMax>>1;

  i = (bytes-1); d = -1; s = this[index+i]; i+=d; nBits = -7;
  for (e = s&((1<<(-nBits))-1), s>>=(-nBits), nBits += eLen; nBits > 0; e=e*256+this[index+i], i+=d, nBits-=8);
  for (m = e&((1<<(-nBits))-1), e>>=(-nBits), nBits += mLen; nBits > 0; m=m*256+this[index+i], i+=d, nBits-=8);

  if (e === 0) {
    // Zero, or denormalized number
    e = 1-eBias;
  } else if (e === eMax) {
    // NaN, or +/-Infinity
    return m?NaN:((s?-1:1)*Infinity);
  } else {
    // Normalized number
    m = m + Math.pow(2, mLen);
    e = e - eBias;
  }
  return (s?-1:1) * m * Math.pow(2, e-mLen);
};

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