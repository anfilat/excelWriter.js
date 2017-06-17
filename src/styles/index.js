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

function Styles(common) {
	this.objectId = common.uniqueId('Styles');
	this.numberFormats = new NumberFormats(this);
	this.fonts = new Fonts(this);
	this.fills = new Fills(this);
	this.borders = new Borders(this);
	this.cells = new Cells(this);
	this.tableElements = new TableElements(this);
	this.tables = new Tables(this);
	this.defaultTableStyle = '';
	this.mergeCache = Object.create(null);
}

Styles.prototype = {
	addFormat(format, name) {
		return this.cells.add(format, name);
	},
	get(name) {
		return this.cells.get(name);
	},
	getId(name) {
		return this.cells.getId(name);
	},
	addFillOutFormat(format) {
		if (this.get(format).fillOut) {
			return format;
		}
		const style = this.cells.cutVisible(this.cells.fullGet(format));

		if (!_.isEmpty(style)) {
			return this.addFormat(style);
		}
	},
	merge(format1, format2, format3) {
		const count = Boolean(format1) + Boolean(format2) + Boolean(format3);

		if (count === 0) {
			return null;
		} else if (count === 1) {
			return this.addFormat(format1 || format2 || format3);
		} else if (count === 2) {
			let f1;
			let f2;

			if (format1) {
				f1 = format1;
				f2 = format2 ? format2 : format3;
			} else {
				f1 = format2;
				f2 = format3;
			}

			return this.merge2(f1, f2);
		} else {
			return this.merge3(format1, format2, format3);
		}
	},
	merge2(format1, format2) {
		const id = JSON.stringify(format1) + '#' + JSON.stringify(format2);
		let merged = this.mergeCache[id];

		if (!merged) {
			let format = {};
			format = this.cells.merge(format, this.cells.fullGet(format1));
			format = this.cells.merge(format, this.cells.fullGet(format2));
			merged = this.cells.add(format, null, {merge: true});
			this.mergeCache[id] = merged;
		}
		return merged;
	},
	merge3(format1, format2, format3) {
		const id = JSON.stringify(format1) + '#' + JSON.stringify(format2) + '#' + JSON.stringify(format3);
		let merged = this.mergeCache[id];

		if (!merged) {
			let format = {};
			format = this.cells.merge(format, this.cells.fullGet(format1));
			format = this.cells.merge(format, this.cells.fullGet(format2));
			format = this.cells.merge(format, this.cells.fullGet(format3));
			merged = this.cells.add(format, null, {merge: true});
			this.mergeCache[id] = merged;
		}
		return merged;
	},
	addFontFormat(format, name) {
		return this.fonts.add(format, name);
	},
	addBorderFormat(format, name) {
		return this.borders.add(format, name);
	},
	addPatternFormat(format, name) {
		return this.fills.add(format, name, {fillType: 'pattern'});
	},
	addGradientFormat(format, name) {
		return this.fills.add(format, name, {fillType: 'gradient'});
	},
	addNumberFormat(format, name) {
		return this.numberFormats.add(format, name);
	},
	addTableFormat(format, name) {
		return this.tables.add(format, name);
	},
	addTableElementFormat(format, name) {
		return this.tableElements.add(format, name);
	},
	setDefaultTableStyle(name) {
		this.tables.defaultTableStyle = name;
	},
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
};

module.exports = Styles;
