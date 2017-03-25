'use strict';

class Images {
	constructor(common) {
		this.common = common;
		this._images = Object.create(null);
		this._imageByNames = Object.create(null);
		this._extensions = Object.create(null);
	}
	addImage(data, type = '', name) {
		let image = this._images[data];

		if (!image) {
			const id = this.common.uniqueIdForSpace('image');
			const path = `xl/media/image${id}.${type}`;
			const contentType = getContentType(type);

			name = name || 'excelWriter' + id;
			image = {
				objectId: 'image' + id,
				data,
				name,
				contentType,
				extension: type,
				path
			};
			this.common.addPath(image, '/' + path);
			this._images[data] = image;
			this._imageByNames[name] = image;
			this._extensions[type] = contentType;
		} else if (name && !this._imageByNames[name]) {
			image.name = name;
			this._imageByNames[name] = image;
		}
		return image.name;
	}
	getImage(name) {
		return this._imageByNames[name];
	}
	getImages() {
		return this._images;
	}
	removeImages() {
		this._images = null;
		this._imageByNames = null;
	}
	getExtensions() {
		return this._extensions;
	}
}

function getContentType(type) {
	switch (type.toLowerCase()) {
		case 'jpeg':
		case 'jpg':
			return 'image/jpeg';
		case 'png':
			return 'image/png';
		case 'gif':
			return 'image/gif';
		default:
			return null;
	}
}

module.exports = Images;
