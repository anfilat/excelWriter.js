'use strict';

const _ = require('lodash');
const util = require('./util');
const Common = require('./common');
const Relations = require('./relations');
const Worksheet = require('./worksheet');
const toXMLString = require('./XMLString');

class Workbook {
	constructor() {
		this.common = new Common();
		this.styles = this.common.styles;
		this.relations = new Relations(this.common);

		this.objectId = this.common.uniqueId('Workbook');
	}
	addWorksheet(config) {
		config = _.defaults(config, {
			name: this.common.getNewWorksheetDefaultName()
		});
		const worksheet = new Worksheet(this, config);
		this.common.addWorksheet(worksheet);
		this.relations.addRelation(worksheet, 'worksheet');

		return worksheet;
	}
	addFormat(format, name) {
		return this.styles.addFormat(format, name);
	}
	addFontFormat(format, name) {
		return this.styles.addFontFormat(format, name);
	}
	addBorderFormat(format, name) {
		return this.styles.addBorderFormat(format, name);
	}
	addPatternFormat(format, name) {
		return this.styles.addPatternFormat(format, name);
	}
	addGradientFormat(format, name) {
		return this.styles.addGradientFormat(format, name);
	}
	addNumberFormat(format, name) {
		return this.styles.addNumberFormat(format, name);
	}
	addTableFormat(format, name) {
		return this.styles.addTableFormat(format, name);
	}
	addTableElementFormat(format, name) {
		return this.styles.addTableElementFormat(format, name);
	}
	setDefaultTableStyle(name) {
		this.styles.setDefaultTableStyle(name);
		return this;
	}
	addImage(data, type, name) {
		return this.common.addImage(data, type, name);
	}
	_generateFiles(zip, canStream) {
		prepareWorksheets(this.common);

		exportWorksheets(zip, canStream, this.common);
		exportTables(zip, this.common);
		exportImages(zip, this.common);
		exportDrawings(zip, this.common);
		exportStyles(zip, this.relations, this.styles);
		exportSharedStrings(zip, canStream, this.relations, this.common);
		zip.file('[Content_Types].xml', createContentTypes(this.common));
		zip.file('_rels/.rels', createWorkbookRelationship());
		zip.file('xl/workbook.xml', this._export());
		zip.file('xl/_rels/workbook.xml.rels', this.relations.export());
	}
	_export() {
		return toXMLString({
			name: 'workbook',
			ns: 'spreadsheetml',
			attributes: [
				['xmlns:r', util.schemas.relationships]
			],
			children: [
				bookViewsXML(this.common),
				sheetsXML(this.relations, this.common),
				definedNamesXML(this.common)
			]
		});
	}
}

function bookViewsXML(common) {
	let activeTab = 0;
	let activeWorksheetId;

	if (common.activeWorksheet) {
		activeWorksheetId = common.activeWorksheet.objectId;

		activeTab = Math.max(activeTab,
			_.findIndex(common.worksheets, worksheet => worksheet.objectId === activeWorksheetId));
	}

	return toXMLString({
		name: 'bookViews',
		children: [
			toXMLString({
				name: 'workbookView',
				attributes: [
					['activeTab', activeTab]
				]
			})
		]
	});
}

function sheetsXML(relations, common) {
	const maxWorksheetNameLength = 31;
	const children = _.map(common.worksheets, (worksheet, index) => {
		// Microsoft Excel (2007, 2013) do not allow worksheet names longer than 31 characters
		// if the worksheet name is longer, Excel displays an 'Excel found unreadable content...' popup when opening the file
		if (worksheet.name.length > maxWorksheetNameLength) {
			throw 'Microsoft Excel requires work sheet names to be less than ' + (maxWorksheetNameLength + 1) +
			' characters long, work sheet name "' + worksheet.name + '" is ' + worksheet.name.length + ' characters long';
		}

		return toXMLString({
			name: 'sheet',
			attributes: [
				['name', worksheet.name],
				['sheetId', index + 1],
				['r:id', relations.getRelationshipId(worksheet)],
				['state', worksheet.getState()]
			]
		});
	});

	return toXMLString({
		name: 'sheets',
		children
	});
}

function definedNamesXML(common) {
	const isPrintTitles = _.some(common.worksheets,
		worksheet => worksheet._printTitles && (worksheet._printTitles.topTo >= 0 || worksheet._printTitles.leftTo >= 0));

	if (isPrintTitles) {
		const children = [];

		_.forEach(common.worksheets, (worksheet, index) => {
			const entry = worksheet._printTitles;

			if (entry && (entry.topTo >= 0 || entry.leftTo >= 0)) {
				const name = worksheet.name;
				let value = '';

				if (entry.topTo >= 0) {
					value = name + '!$' + (entry.topFrom + 1) + ':$' + (entry.topTo + 1);
					if (entry.leftTo >= 0) {
						value += ',';
					}
				}
				if (entry.leftTo >= 0) {
					value += name + '!$' + util.positionToLetter(entry.leftFrom + 1) +
						':$' + util.positionToLetter(entry.leftTo + 1);
				}

				children.push(toXMLString({
					name: 'definedName',
					value,
					attributes: [
						['name', '_xlnm.Print_Titles'],
						['localSheetId', index]
					]
				}));
			}
		});

		return toXMLString({
			name: 'definedNames',
			children
		});
	}
	return '';
}

function prepareWorksheets(common) {
	_.forEach(common.worksheets, worksheet => {
		worksheet._prepare();
	});
}

function exportWorksheets(zip, canStream, common) {
	_.forEach(common.worksheets, worksheet => {
		zip.file(worksheet.path, worksheet._export(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations.export());
	});
}

function exportTables(zip, common) {
	_.forEach(common.tables, table => {
		zip.file(table.path, table._export());
	});
}

function exportImages(zip, common) {
	_.forEach(common.getImages(), image => {
		zip.file(image.path, image.data, {base64: true, binary: true});
		image.data = null;
	});
	common.removeImages();
}

function exportDrawings(zip, common) {
	_.forEach(common.drawings, drawing => {
		zip.file(drawing.path, drawing.export());
		zip.file(drawing.relationsPath, drawing.relations.export());
	});
}

function exportStyles(zip, relations, styles) {
	relations.addRelation(styles, 'stylesheet');
	zip.file('xl/styles.xml', styles.export());
}

function exportSharedStrings(zip, canStream, relations, common) {
	if (!common.sharedStrings.isEmpty()) {
		relations.addRelation(common.sharedStrings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.sharedStrings.export(canStream));
	}
}

function createContentTypes(common) {
	const children = [];

	children.push(toXMLString({
		name: 'Default',
		attributes: [
			['Extension', 'rels'],
			['ContentType', 'application/vnd.openxmlformats-package.relationships+xml']
		]
	}));
	children.push(toXMLString({
		name: 'Default',
		attributes: [
			['Extension', 'xml'],
			['ContentType', 'application/xml']
		]
	}));
	children.push(toXMLString({
		name: 'Override',
		attributes: [
			['PartName', '/xl/workbook.xml'],
			['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']
		]
	}));
	if (!common.sharedStrings.isEmpty()) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/sharedStrings.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml']
			]
		}));
	}
	children.push(toXMLString({
		name: 'Override',
		attributes: [
			['PartName', '/xl/styles.xml'],
			['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']
		]
	}));

	_.forEach(common.worksheets, (worksheet, index) => {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
			]
		}));
	});
	_.forEach(common.tables, (table, index) => {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/tables/table' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']
			]
		}));
	});
	_.forEach(common.getExtensions(), (contentType, extension) => {
		children.push(toXMLString({
			name: 'Default',
			attributes: [
				['Extension', extension],
				['ContentType', contentType]
			]
		}));
	});
	_.forEach(common.drawings, (drawing, index) => {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/drawings/drawing' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml']
			]
		}));
	});

	return toXMLString({
		name: 'Types',
		ns: 'contentTypes',
		children
	});
}

function createWorkbookRelationship() {
	return toXMLString({
		name: 'Relationships',
		ns: 'relationshipPackage',
		children: [
			toXMLString({
				name: 'Relationship',
				attributes: [
					['Id', 'rId1'],
					['Type', util.schemas.officeDocument],
					['Target', 'xl/workbook.xml']
				]
			})
		]
	});
}

module.exports = Workbook;
