'use strict';

function Paths() {
	this.paths = {};
}

Paths.prototype.add = function (object, path) {
	this.paths[object.objectId] = path;
};

Paths.prototype.get = function (object) {
	return this.paths[object.objectId];
};

module.exports = Paths;
