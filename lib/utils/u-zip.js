const events = require('events');
const UZIP = require('uzip');
const StreamBuf = require('./stream-buf');
const {stringToBuffer} = require('./browser-buffer-encode');

class UZipWriter extends events.EventEmitter {
  constructor() {
    super();
    this.files = {};
    this.stream = new StreamBuf();
  }

  append(data, options) {
    if(typeof data === 'string'){
        this.files[options.name] =  stringToBuffer(data);
    }
    else{
        this.files[options.name] =  new Uint8Array(data);
    }
     
  }

  async finalize(){
     const content=  UZIP.encode(this.files);
    this.stream.end(content);
    this.emit('finish');
  }

  pipe(destination, options) {
    return this.stream.pipe(destination, options);
  }
}


module.exports = {
    UZipWriter,
  };

