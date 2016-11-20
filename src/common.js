'use strict';

var _ = require('lodash');
var Paths = require('./paths');
var SharedStrings = require('./sharedStrings');
var Styles = require('./styles');
var util = require('./util');

function Common() {
	this.paths = new Paths();

	this.sharedStrings = new SharedStrings();
	this.paths.add(this.sharedStrings, 'sharedStrings.xml');

	this.styles = new Styles();
	this.paths.add(this.styles, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.images = {};
	this.imageByNames = {};
	this.drawings = [];
}

Common.prototype.addWorksheet = function (worksheet) {
	var index = this.worksheets.length + 1;
	var path = 'worksheets/sheet' + index + '.xml';
	var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

	worksheet.path = 'xl/' + path;
	worksheet.relationsPath = relationsPath;
	this.worksheets.push(worksheet);
	this.paths.add(worksheet, path);
};

Common.prototype.getNewWorksheetDefaultName = function () {
	return 'Sheet ' + (this.worksheets.length + 1);
};

Common.prototype.setActiveWorksheet = function (worksheet) {
	this.activeWorksheet = worksheet;
};

Common.prototype.addTable = function (table) {
	var index = this.tables.length + 1;
	var path = 'xl/tables/table' + index + '.xml';

	table.path = path;
	this.tables.push(table);
	this.paths.add(table, '/' + path);
};

Common.prototype.addImage = function (data, type, name) {
	var image = this.images[data];

	if (!image) {
		type = type || '';

		var contentType;
		var id = util.uniqueId('image');
		var path = 'xl/media/image' + id + '.' + type;

		name = name || '_jsExcelWriter' + id;
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
			objectId: _.uniqueId('Image'),
			data: data,
			name: name,
			contentType: contentType,
			extension: type,
			path: path
		};
		this.paths.add(image, '/' + path);
		this.images[data] = image;
		this.imageByNames[name] = image;
	} else if (name && !this.imageByNames[name]) {
		image.name = name;
		this.imageByNames[name] = image;
	}
	return image.name;
};

Common.prototype.getImage = function (name) {
	return this.imageByNames[name];
};

Common.prototype.addDrawings = function (drawings) {
	var index = this.drawings.length + 1;
	var path = 'xl/drawings/drawing' + index + '.xml';
	var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

	drawings.path = path;
	drawings.relationsPath = relationsPath;
	this.drawings.push(drawings);
	this.paths.add(drawings, '/' + path);
};

module.exports = Common;
