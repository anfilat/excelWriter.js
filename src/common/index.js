'use strict';

const Images = require('./images');
const SharedStrings = require('./sharedStrings');
const Styles = require('../styles');

function Common() {
	this.idSpaces = Object.create(null);
	this.paths = Object.create(null);

	this.images = new Images(this);

	this.strings = new SharedStrings(this);
	this.addPath(this.strings, 'sharedStrings.xml');

	this.styles = new Styles(this);
	this.addPath(this.styles, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.drawings = [];
}

Common.prototype = {
	uniqueId(space) {
		return space + this.uniqueIdForSpace(space);
	},
	uniqueIdForSpace(space) {
		if (!this.idSpaces[space]) {
			this.idSpaces[space] = 1;
		}
		return this.idSpaces[space]++;
	},

	addPath(object, path) {
		this.paths[object.objectId] = path;
	},
	getPath(object) {
		return this.paths[object.objectId];
	},

	addWorksheet(worksheet) {
		const index = this.worksheets.length + 1;
		const path = 'worksheets/sheet' + index + '.xml';
		const relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

		worksheet.path = 'xl/' + path;
		worksheet.relationsPath = relationsPath;
		this.worksheets.push(worksheet);
		this.addPath(worksheet, path);
	},
	getNewWorksheetDefaultName() {
		return `Sheet ${this.worksheets.length + 1}`;
	},
	setActiveWorksheet(worksheet) {
		this.activeWorksheet = worksheet;
	},
	getActiveWorksheetIndex() {
		if (this.activeWorksheet) {
			const activeWorksheetId = this.activeWorksheet.objectId;

			return Math.max(0, this.worksheets.findIndex(worksheet => worksheet.objectId === activeWorksheetId));
		}
		return 0;
	},

	addTable(table) {
		const index = this.tables.length + 1;
		const path = 'xl/tables/table' + index + '.xml';

		table.path = path;
		this.tables.push(table);
		this.addPath(table, '/' + path);
	},

	addDrawings(drawings) {
		const index = this.drawings.length + 1;
		const path = 'xl/drawings/drawing' + index + '.xml';
		const relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

		drawings.path = path;
		drawings.relationsPath = relationsPath;
		this.drawings.push(drawings);
		this.addPath(drawings, '/' + path);
	}
};

module.exports = Common;
