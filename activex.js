var activex = module.exports = require('./build/Debug/node_activex');

global.ActiveXObject = function(id) {
    return new activex.Object(id);
};
