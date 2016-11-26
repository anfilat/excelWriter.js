'use strict';

module.exports = {
	init: function () {
		this._images = Object.create(null);
		this._imageByNames = Object.create(null);
		this._extensions = Object.create(null);
	},
	methods: {
		addImage: function (data, type, name) {
			var image = this._images[data];

			if (!image) {
				type = type || '';

				var contentType;
				var id = this.uniqueIdSeparated('image').id;
				var path = 'xl/media/image' + id + '.' + type;

				name = name || 'excelWriter' + id;
				switch (type.toLowerCase()) {
					case 'jpeg':
					case 'jpg':
						contentType = 'image/jpeg';
						break;
					case 'png':
						contentType = 'image/png';
						break;
					case 'gif':
						contentType = 'image/gif';
						break;
					default:
						contentType = null;
						break;
				}

				image = {
					objectId: 'image' + id,
					data: data,
					name: name,
					contentType: contentType,
					extension: type,
					path: path
				};
				this.addPath(image, '/' + path);
				this._images[data] = image;
				this._imageByNames[name] = image;
				this._extensions[type] = contentType;
			} else if (name && !this._imageByNames[name]) {
				image.name = name;
				this._imageByNames[name] = image;
			}
			return image.name;
		},
		getImage: function (name) {
			return this._imageByNames[name];
		},
		getImages: function () {
			return this._images;
		},
		removeImages: function () {
			this._images = null;
			this._imageByNames = null;
		},
		getExtensions: function () {
			return this._extensions;
		}
	}
};
