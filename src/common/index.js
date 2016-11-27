'use strict';

var _ = require('lodash');
var paths = require('./paths');
var images = require('./images');
var SharedStrings = require('./sharedStrings');
var Styles = require('../styles');

function Common() {
	paths.init.call(this);
	images.init.call(this);

	this.idSpaces = Object.create(null);

	this.sharedStrings = new SharedStrings(this);
	this.addPath(this.sharedStrings, 'sharedStrings.xml');

	this.styles = new Styles(this);
	this.addPath(this.styles, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.drawings = [];
}

_.assign(Common.prototype, paths.methods);
_.assign(Common.prototype, images.methods);

Common.prototype.uniqueId = function (space) {
	if (!this.idSpaces[space]) {
		this.idSpaces[space] = 1;
	}
	return space + this.idSpaces[space]++;
};

Common.prototype.uniqueIdSeparated = function (space) {
	if (!this.idSpaces[space]) {
		this.idSpaces[space] = 1;
	}
	return {
		space: space,
		id: this.idSpaces[space]++
	};
};

Common.prototype.addString = function (string) {
	return this.sharedStrings.add(string);
};

Common.prototype.addWorksheet = function (worksheet) {
	var index = this.worksheets.length + 1;
	var path = 'worksheets/sheet' + index + '.xml';
	var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

	worksheet.path = 'xl/' + path;
	worksheet.relationsPath = relationsPath;
	this.worksheets.push(worksheet);
	this.addPath(worksheet, path);
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
	this.addPath(table, '/' + path);
};

Common.prototype.addDrawings = function (drawings) {
	var index = this.drawings.length + 1;
	var path = 'xl/drawings/drawing' + index + '.xml';
	var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

	drawings.path = path;
	drawings.relationsPath = relationsPath;
	this.drawings.push(drawings);
	this.addPath(drawings, '/' + path);
};

module.exports = Common;
