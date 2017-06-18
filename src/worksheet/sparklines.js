'use strict';

const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

const defaultTypeAttrs = {
	displayEmptyCellsAs: 'gap'
};

const defaultTypeColors = {
	color: 'FF376092',
	colorNegative: 'FFD00000',
	colorAxis: 'FF000000',
	colorMarkers: 'FFD00000',
	colorFirst: 'FFD00000',
	colorLast: 'FFD00000',
	colorHigh: 'FFD00000',
	colorLow: 'FFD00000'
};

function Sparklines(worksheet) {
	this.worksheet = worksheet;
	this.lines = Object.create(null);
	this.defaultLines = {
		type: null,
		lines: []
	};
	this.typesByData = Object.create(null);
	this.lastId = 1;
}

Sparklines.prototype = {
	addType(params, name) {
		const type = this.canonType(params);
		const stringType = JSON.stringify(type);

		if (!name && this.typesByData[stringType]) {
			return this.typesByData[stringType];
		}
		return this.addNew(type, stringType, name);
	},
	setDefaultType(params) {
		this.defaultLines.type = this.canonType(params);
	},
	add({cell, beginCell, endCell, type} = {}) {
		const line = {
			cell: util.canonCell(cell),
			beginCell: util.canonCell(beginCell),
			endCell: util.canonCell(endCell)
		};

		if (type) {
			this.lines(type).lines.push(line);
		} else {
			this.defaultLines.lines.push(line);
		}
	},
	insert() {
	},

	canonType(params = {}) {
		const type = {};

		_.keys(defaultTypeAttrs)
			.forEach(property => {
				type[property] = params[property] || defaultTypeAttrs[property];
			});
		_.keys(defaultTypeColors)
			.forEach(property => {
				type[property] = util.canonColor(params[property] || defaultTypeColors[property]);
			});
		return type;
	},
	addNew(type, stringType, name) {
		name = name || 'Sparklines' + this.lastId++;

		this.typesByData[stringType] = name;
		this.lines(type).type = type;

		return name;
	},
	lines(type) {
		this.lines[type] = this.lines[type] || {
			type: null,
			lines: []
		};
		return this.lines[type];
	},

	save() {
		const children = [];

		if (this.defaultLines.lines.length) {
			const type = this.defaultLines.type || this.canonType();

			children.push(this.saveType(type, this.defaultLines.lines));
		}

		_.forEach(this.lines, ({type, lines}) => {
			if (type && lines.length) {
				children.push(this.saveType(type, lines));
			}
		});

		if (children.length) {
			return toXMLString({
				name: 'ext',
				attributes: [
					['uri', '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'],
					['xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main']
				],
				children: [
					toXMLString({
						name: 'x14:sparklineGroups',
						attributes : [
							['xmlns:xm', 'http://schemas.microsoft.com/office/excel/2006/main']
						],
						children
					})
				]
			});
		}
		return '';
	},
	saveType(type, lines) {
		const attributes = [];
		const children = [];

		_.keys(defaultTypeAttrs)
			.forEach(property => {
				attributes.push([property, type[property]]);
			});
		_.keys(defaultTypeColors)
			.forEach(property => {
				const name = property === 'color' ? 'colorSeries' : property;
				const value = type[property];

				children.push(`<x14:${name} rgb="${value}"/>`);
			});
		children.push(toXMLString({
			name : 'x14:sparklines',
			children: this.saveLines(lines)
		}));

		return toXMLString({
			name : 'x14:sparklineGroup',
			attributes,
			children
		});
	},
	saveLines(lines) {
		return lines.map(line => toXMLString({
			name: 'x14:sparkline',
			children: [
				`<xm:f>${this.worksheet.sheetName()}!${line.beginCell}:${line.endCell}</xm:f>`,
				`<xm:sqref>${line.cell}</xm:sqref>`
			]
		}));
	}
};

module.exports = Sparklines;
