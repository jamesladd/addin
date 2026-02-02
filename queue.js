/* Simple Queue Implementation for InboxAgent */

(function (global) {
  'use strict';

  /**
   * Queue constructor
   * @param {Object} options - Configuration options
   * @param {number} options.concurrency - Number of concurrent jobs (default: Infinity)
   * @param {number} options.timeout - Default timeout for jobs in ms (default: 0 = no timeout)
   * @param {Array} options.results - Array to store results
   */
  function Queue(options) {
    options = options || {};

    this.concurrency = options.concurrency || Infinity;
    this.timeout = options.timeout || 0;
    this.results = options.results || [];
    this.pending = 0;
    this.jobs = [];
    this.running = false;
    this.ended = false;
    this.listeners = {};
  }

  /**
   * Add event listener
   */
  Queue.prototype.addEventListener = function (event, callback) {
    if (!this.listeners[event]) {
      this.listeners[event] = [];
    }
    this.listeners[event].push(callback);
  };

  /**
   * Remove event listener
   */
  Queue.prototype.removeEventListener = function (event, callback) {
    if (!this.listeners[event]) return;
    const index = this.listeners[event].indexOf(callback);
    if (index > -1) {
      this.listeners[event].splice(index, 1);
    }
  };

  /**
   * Emit event
   */
  Queue.prototype.emit = function (event, detail) {
    if (!this.listeners[event]) return;
    const eventObj = { detail: detail, type: event };
    this.listeners[event].forEach(function (callback) {
      callback(eventObj);
    });
  };

  /**
   * Add job to end of queue
   */
  Queue.prototype.push = function () {
    const self = this;
    for (let i = 0; i < arguments.length; i++) {
      self.jobs.push(arguments[i]);
    }
    if (self.running) {
      self.process();
    }
    return self;
  };

  /**
   * Add job to beginning of queue
   */
  Queue.prototype.unshift = function () {
    const self = this;
    for (let i = arguments.length - 1; i >= 0; i--) {
      self.jobs.unshift(arguments[i]);
    }
    if (self.running) {
      self.process();
    }
    return self;
  };

  /**
   * Splice jobs into queue
   */
  Queue.prototype.splice = function (start, deleteCount) {
    const self = this;
    const args = Array.prototype.slice.call(arguments, 2);
    Array.prototype.splice.apply(self.jobs, [start, deleteCount].concat(args));
    if (self.running) {
      self.process();
    }
    return self;
  };

  /**
   * Start processing queue
   */
  Queue.prototype.start = function (callback) {
    const self = this;

    if (self.running) return self;

    self.running = true;
    self.ended = false;
    self.callback = callback || function (err) {
      if (err) throw err;
    };

    self.process();
    return self;
  };

  /**
   * Stop processing queue
   */
  Queue.prototype.stop = function () {
    this.running = false;
    return this;
  };

  /**
   * End queue
   */
  Queue.prototype.end = function (err) {
    const self = this;

    if (self.ended) return self;

    self.ended = true;
    self.running = false;

    if (self.callback) {
      self.callback(err, self.results);
    }

    return self;
  };

  /**
   * Process next job
   */
  Queue.prototype.process = function () {
    const self = this;

    // Check if we should continue
    if (!self.running) return;
    if (self.pending >= self.concurrency) return;
    if (self.jobs.length === 0) {
      if (self.pending === 0) {
        self.end();
      }
      return;
    }

    // Get next job
    const job = self.jobs.shift();
    self.pending++;

    // Determine timeout for this job
    const timeout = typeof job.timeout !== 'undefined' ? job.timeout : self.timeout;

    let timeoutId = null;
    let timedOut = false;
    let completed = false;

    // Create callback wrapper
    const next = function (err, result) {
      if (completed) return;
      completed = true;

      if (timeoutId) {
        clearTimeout(timeoutId);
      }

      self.pending--;

      if (!timedOut) {
        if (!err && typeof result !== 'undefined') {
          self.results.push(result);
        }

        self.emit('success', { job: job, result: result });

        if (err) {
          self.emit('error', { job: job, error: err });
          self.end(err);
          return;
        }
      }

      self.process();
    };

    // Set timeout if specified
    if (timeout > 0) {
      timeoutId = setTimeout(function () {
        if (completed) return;
        timedOut = true;
        completed = true;

        self.pending--;

        self.emit('timeout', {
          job: job,
          next: function () {
            self.process();
          }
        });
      }, timeout);
    }

    // Execute job
    try {
      const result = job(next);

      // Handle promises
      if (result && typeof result.then === 'function') {
        result
          .then(function (res) {
            next(null, res);
          })
          .catch(function (err) {
            next(err);
          });
      }
    } catch (err) {
      next(err);
    }
  };

  // Export Queue
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = Queue;
  } else {
    global.Queue = Queue;
  }

})(typeof window !== 'undefined' ? window : global);