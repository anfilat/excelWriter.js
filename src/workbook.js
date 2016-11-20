'use strict';

var _ = require('lodash');
var util = require('./util');
var Common = require('./common');
var RelationshipManager = require('./relationshipManager');
var Worksheet = require('./worksheet');
var toXMLString = require('./XMLString');

function Workbook() {
	this.objectId = _.uniqueId('Workbook');

	this.common = new Common();
	this.relations = new RelationshipManager(this.common);
	this.styles = this.common.styles;
}

Workbook.prototype.addWorksheet = function (config) {
	var worksheet;

	config = _.defaults(config, {
		name: this.common.getNewWorksheetDefaultName()
	});
	worksheet = new Worksheet(this, config);
	this.common.addWorksheet(worksheet);
	this.relations.addRelation(worksheet, 'worksheet');

	return worksheet;
};

Workbook.prototype.addFormat = function (format, name) {
	return this.styles.addFormat(format, name);
};

Workbook.prototype.addFontFormat = function (format, name) {
	return this.styles.addFontFormat(format, name);
};

Workbook.prototype.addBorderFormat = function (format, name) {
	return this.styles.addBorderFormat(format, name);
};

Workbook.prototype.addPatternFormat = function (format, name) {
	return this.styles.addPatternFormat(format, name);
};

Workbook.prototype.addGradientFormat = function (format, name) {
	return this.styles.addGradientFormat(format, name);
};

Workbook.prototype.addNumberFormat = function (format, name) {
	return this.styles.addNumberFormat(format, name);
};

Workbook.prototype.addTableFormat = function (format, name) {
	return this.styles.addTableFormat(format, name);
};

Workbook.prototype.addTableElementFormat = function (format, name) {
	return this.styles.addTableElementFormat(format, name);
};

Workbook.prototype.setDefaultTableStyle = function (name) {
	this.styles.setDefaultTableStyle(name);
	return this;
};

Workbook.prototype.addImage = function (data, type, name) {
	return this.common.addImage(data, type, name);
};

Workbook.prototype._generateFiles = function (zip, canStream) {
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
	zip.file('xl/_rels/workbook.xml.rels', this.relations._export());
};

Workbook.prototype._export = function () {
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
};

function bookViewsXML(common) {
	var activeTab = 0;
	var activeWorksheetId;

	if (common.activeWorksheet) {
		activeWorksheetId = common.activeWorksheet.objectId;

		activeTab = Math.max(activeTab, _.findIndex(common.worksheets, function (worksheet) {
			return worksheet.objectId === activeWorksheetId;
		}));
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
	var maxWorksheetNameLength = 31;
	var children = _.map(common.worksheets, function (worksheet, index) {
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
				['state', worksheet.state]
			]
		});
	});

	return toXMLString({
		name: 'sheets',
		children: children
	});
}

function definedNamesXML(common) {
	var isPrintTitles = _.some(common.worksheets, function (worksheet) {
		return worksheet._printTitles && (worksheet._printTitles.topTo >= 0 || worksheet._printTitles.leftTo >= 0);
	});

	if (isPrintTitles) {
		var children = [];

		_.forEach(common.worksheets, function (worksheet, index) {
			var entry = worksheet._printTitles;

			if (entry && (entry.topTo >= 0 || entry.leftTo >= 0)) {
				var value = '';
				var name = worksheet.name;

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
					value: value,
					attributes: [
						['name', '_xlnm.Print_Titles'],
						['localSheetId', index]
					]
				}));
			}
		});

		return toXMLString({
			name: 'definedNames',
			children: children
		});
	}
	return '';
}

function prepareWorksheets(common) {
	_.forEach(common.worksheets, function (worksheet) {
		worksheet._prepare();
	});
}

function exportWorksheets(zip, canStream, common) {
	_.forEach(common.worksheets, function (worksheet) {
		zip.file(worksheet.path, worksheet._export(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations._export());
	});
}

function exportTables(zip, common) {
	_.forEach(common.tables, function (table) {
		zip.file(table.path, table._export());
	});
}

function exportImages(zip, common) {
	_.forEach(common.images, function (image) {
		zip.file(image.path, image.data, {base64: true, binary: true});
		image.data = null;
	});
	common.images = null;
}

function exportDrawings(zip, common) {
	_.forEach(common.drawings, function (drawing) {
		zip.file(drawing.path, drawing._export());
		zip.file(drawing.relationsPath, drawing.relations._export());
	});
}

function exportStyles(zip, relations, styles) {
	relations.addRelation(styles, 'stylesheet');
	zip.file('xl/styles.xml', styles._export());
}

function exportSharedStrings(zip, canStream, relations, common) {
	if (!common.sharedStrings.isEmpty()) {
		relations.addRelation(common.sharedStrings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.sharedStrings._export(canStream));
	}
}

function createContentTypes(common) {
	var children = [];
	var extensions = {};

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

	_.forEach(common.worksheets, function (worksheet, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
			]
		}));
	});
	_.forEach(common.tables, function (table, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/tables/table' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']
			]
		}));
	});
	_.forEach(common.imageByNames, function (image) {
		extensions[image.extension] = image.contentType;
	});
	_.forEach(extensions, function (contentType, extension) {
		children.push(toXMLString({
			name: 'Default',
			attributes: [
				['Extension', extension],
				['ContentType', contentType]
			]
		}));
	});
	_.forEach(common.drawings, function (drawing, index) {
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
		children: children
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
