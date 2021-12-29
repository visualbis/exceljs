const events = require('events');
const clientZip = require('client-zip');
const StreamBuf = require('./stream-buf');

class ClientZipWriter extends events.EventEmitter {
  constructor() {
    super();
    this.files = [];
    this.stream = new StreamBuf();
  }

  append(data, options) {

    const currFile = {
        name: options.name,
        lastModified: new Date(),
        input: data,
    };
    this.files.push(currFile);
    
  }

  async finalize(){
    const content=  await clientZip.downloadZip(this.files).text();
    this.stream.end(content);
    this.emit('finish');
  }

  pipe(destination, options) {
    return this.stream.pipe(destination, options);
  }
}


module.exports = {
    ClientZipWriter,
  };

