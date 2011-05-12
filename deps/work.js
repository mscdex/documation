var Work = module.exports = function(tasks, cb) {
  this._tasks = (!tasks || !Array.isArray(tasks) ? new Array() : tasks);
  this._cb = (typeof cb === 'undefined' ? tasks : cb);
};
Work.prototype.push = function(fn) {
  this._tasks.push(fn);
};
Work.prototype.go = function() {
  this._task = 0;
  this.next();
};
Work.prototype.next = function(err) {
  if (err)
    return this._cb(err);
  else {
    var nextCb = (this._task < this._tasks.length ?
                  this._tasks[this._task++] : this._cb);
    if (nextCb)
      process.nextTick(nextCb);
  }
};