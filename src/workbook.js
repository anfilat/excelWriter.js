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
		this.images = this.common.images;

		this.objectId = this.common.uniqueId('Workbook');

		this.relations = new Relations(this.common);
		this.relations.addRelation(this.styles, 'stylesheet');
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
		return this.images.addImage(data, type, name);
	}
	_generateFiles(zip, canStream) {
		prepareWorksheets(this.common);

		saveWorksheets(zip, canStream, this.common);
		saveTables(zip, this.common);
		saveImages(zip, this.images);
		saveDrawings(zip, this.common);
		saveStyles(zip, this.styles);
		saveStrings(zip, canStream, this.relations, this.common);
		zip.file('[Content_Types].xml', createContentTypes(this.common));
		zip.file('_rels/.rels', createWorkbookRelationship());
		zip.file('xl/workbook.xml', this._save());
		zip.file('xl/_rels/workbook.xml.rels', this.relations.save());
	}
	_save() {
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

	if (common.activeWorksheet) {
		const activeWorksheetId = common.activeWorksheet.objectId;

		activeTab = Math.max(activeTab,
			common.worksheets.findIndex(worksheet => worksheet.objectId === activeWorksheetId));
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
	const children = common.worksheets.map((worksheet, index) => {
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
	const isPrintTitles = common.worksheets.some(
		worksheet => worksheet._printTitles && (worksheet._printTitles.topTo >= 0 || worksheet._printTitles.leftTo >= 0)
	);

	if (isPrintTitles) {
		const children = [];

		common.worksheets.forEach((worksheet, index) => {
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
	common.worksheets.forEach(worksheet => {
		worksheet._prepare();
	});
}

function saveWorksheets(zip, canStream, common) {
	common.worksheets.forEach(worksheet => {
		zip.file(worksheet.path, worksheet._save(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations.save());
	});
}

function saveTables(zip, common) {
	common.tables.forEach(table => {
		zip.file(table.path, table._save());
	});
}

function saveImages(zip, images) {
	_.forEach(images.getImages(), image => {
		zip.file(image.path, image.data, {base64: true, binary: true});
		image.data = null;
	});
	images.removeImages();
}

function saveDrawings(zip, common) {
	common.drawings.forEach(drawing => {
		zip.file(drawing.path, drawing.save());
		zip.file(drawing.relationsPath, drawing.relations.save());
	});
}

function saveStyles(zip, styles) {
	zip.file('xl/styles.xml', styles.save());
}

function saveStrings(zip, canStream, relations, common) {
	if (common.strings.isStrings()) {
		relations.addRelation(common.strings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.strings.save(canStream));
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
	if (common.strings.isStrings()) {
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

	common.worksheets.forEach((worksheet, index) => {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
			]
		}));
	});
	common.tables.forEach((table, index) => {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/tables/table' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']
			]
		}));
	});
	_.forEach(common.images.getExtensions(), (contentType, extension) => {
		children.push(toXMLString({
			name: 'Default',
			attributes: [
				['Extension', extension],
				['ContentType', contentType]
			]
		}));
	});
	common.drawings.forEach((drawing, index) => {
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
