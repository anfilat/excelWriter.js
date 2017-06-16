'use strict';

function Images(common) {
	this.common = common;
	this.images = Object.create(null);
	this.imageByNames = Object.create(null);
	this.extensions = Object.create(null);
}

Images.prototype = {
	add(data, type = '', name) {
		let image = this.images[data];

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
			this.images[data] = image;
			this.imageByNames[name] = image;
			this.extensions[type] = contentType;
		} else if (name && !this.imageByNames[name]) {
			image.name = name;
			this.imageByNames[name] = image;
		}
		return image.name;
	},
	get(name) {
		return this.imageByNames[name];
	},
	getImages() {
		return this.images;
	},
	removeImages() {
		this.images = null;
		this.imageByNames = null;
	},
	getExtensions() {
		return this.extensions;
	}
};

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
