'use strict';

module.exports = {
	init: function () {
		this._paths = Object.create(null);
	},
	methods: {
		addPath: function (object, path) {
			this._paths[object.objectId] = path;
		},
		getPath: function (object) {
			return this._paths[object.objectId];
		}
	}
};
