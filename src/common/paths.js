'use strict';

module.exports = SuperClass => class Paths extends SuperClass {
	constructor() {
		super();
		this._paths = new Map();
	}
	addPath(object, path) {
		this._paths.set(object.objectId, path);
	}
	getPath(object) {
		return this._paths.get(object.objectId);
	}
};
