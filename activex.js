var ActiveX = module.exports = require('./build/Release/node_activex');

global.ActiveXObject = function(id) {
    return new ActiveX.Object(id);
};
