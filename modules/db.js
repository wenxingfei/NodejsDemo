/**
 * Created by wWX247609 on 2016/4/12.
 */
var Promise = require("bluebird");
var mongoskin = require("mongoskin");

Object.keys(mongoskin).forEach(function(key) {
    var value = mongoskin[key];
    if (typeof value === "function") {
        Promise.promisifyAll(value);
        Promise.promisifyAll(value.prototype);
    }
});
Promise.promisifyAll(mongoskin);

var db = mongoskin.db('mongodb://localhost/NodeJSDemo', { native_parser: true });
module.exports = db;