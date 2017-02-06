'use strict';

class Paths {
	constructor() {
		this._paths = new Map();
	}
	addPath(object, path) {
		this._paths.set(object.objectId, path);
	}
	getPath(object) {
		return this._paths.get(object.objectId);
	}
}

module.exports = Paths;
