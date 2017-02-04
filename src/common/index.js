'use strict';

const paths = require('./paths');
const images = require('./images');
const SharedStrings = require('./sharedStrings');
const Styles = require('../styles');

class Common extends images(paths(Object)) {
	constructor() {
		super();

		this.idSpaces = Object.create(null);

		this.sharedStrings = new SharedStrings({common: this});
		this.addPath(this.sharedStrings, 'sharedStrings.xml');

		this.styles = new Styles(this);
		this.addPath(this.styles, 'styles.xml');

		this.worksheets = [];
		this.tables = [];
		this.drawings = [];
	}
	uniqueId(space) {
		return space + this.uniqueIdForSpace(space);
	}
	uniqueIdForSpace(space) {
		if (!this.idSpaces[space]) {
			this.idSpaces[space] = 1;
		}
		return this.idSpaces[space]++;
	}
	addString(string) {
		return this.sharedStrings.add(string);
	}
	addWorksheet(worksheet) {
		const index = this.worksheets.length + 1;
		const path = 'worksheets/sheet' + index + '.xml';
		const relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

		worksheet.path = 'xl/' + path;
		worksheet.relationsPath = relationsPath;
		this.worksheets.push(worksheet);
		this.addPath(worksheet, path);
	}
	getNewWorksheetDefaultName() {
		return `Sheet ${this.worksheets.length + 1}`;
	}
	setActiveWorksheet(worksheet) {
		this.activeWorksheet = worksheet;
	}
	addTable(table) {
		const index = this.tables.length + 1;
		const path = 'xl/tables/table' + index + '.xml';

		table.path = path;
		this.tables.push(table);
		this.addPath(table, '/' + path);
	}
	addDrawings(drawings) {
		const index = this.drawings.length + 1;
		const path = 'xl/drawings/drawing' + index + '.xml';
		const relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

		drawings.path = path;
		drawings.relationsPath = relationsPath;
		this.drawings.push(drawings);
		this.addPath(drawings, '/' + path);
	}
}

module.exports = Common;
