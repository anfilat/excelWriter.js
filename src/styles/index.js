'use strict';

const _ = require('lodash');
const NumberFormats = require('./numberFormats');
const Fonts = require('./fonts');
const Fills = require('./fills');
const Borders = require('./borders');
const Cells = require('./cells');
const Tables = require('./tables');
const TableElements = require('./tableElements');
const toXMLString = require('../XMLString');

class Styles {
	constructor(common) {
		this.objectId = common.uniqueId('Styles');
		this.numberFormats = new NumberFormats(this);
		this.fonts = new Fonts(this);
		this.fills = new Fills(this);
		this.borders = new Borders(this);
		this.cells = new Cells(this);
		this.tableElements = new TableElements(this);
		this.tables = new Tables(this);
		this.defaultTableStyle = '';
	}
	addFormat(format, name) {
		return this.cells.add(format, name);
	}
	_get(name) {
		return this.cells.get(name);
	}
	_getId(name) {
		return this.cells.getId(name);
	}
	_addInvisibleFormat(format) {
		const style = this.cells.cutVisible(this.cells.fullGet(format));

		if (!_.isEmpty(style)) {
			return this.addFormat(style);
		}
	}
	_merge(format1, format2, format3) {
		const count = Number(Boolean(format1)) + Number(Boolean(format2)) + Number(Boolean(format3));

		if (count === 0) {
			return null;
		} else if (count === 1) {
			return this.cells.add(format1 || format2 || format3);
		} else {
			let format = {};

			if (format1) {
				format = this.cells.merge(format, this.cells.fullGet(format1));
			}
			if (format2) {
				format = this.cells.merge(format, this.cells.fullGet(format2));
			}
			if (format3) {
				format = this.cells.merge(format, this.cells.fullGet(format3));
			}
			return this.cells.add(format, null, {merge: true});
		}
	}
	addFontFormat(format, name) {
		return this.fonts.add(format, name);
	}
	addBorderFormat(format, name) {
		return this.borders.add(format, name);
	}
	addPatternFormat(format, name) {
		return this.fills.add(format, name, {fillType: 'pattern'});
	}
	addGradientFormat(format, name) {
		return this.fills.add(format, name, {fillType: 'gradient'});
	}
	addNumberFormat(format, name) {
		return this.numberFormats.add(format, name);
	}
	addTableFormat(format, name) {
		return this.tables.add(format, name);
	}
	addTableElementFormat(format, name) {
		return this.tableElements.add(format, name);
	}
	setDefaultTableStyle(name) {
		this.tables.defaultTableStyle = name;
	}
	save() {
		return toXMLString({
			name: 'styleSheet',
			ns: 'spreadsheetml',
			children: [
				this.numberFormats.save(),
				this.fonts.save(),
				this.fills.save(),
				this.borders.save(),
				this.cells.save(),
				this.tableElements.save(),
				this.tables.save()
			]
		});
	}
}

module.exports = Styles;
