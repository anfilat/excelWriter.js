'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const JSZip = require('jszip');
const util = require('./util');
const Common = require('./common');
const Relations = require('./relations');
const createWorksheet = require('./worksheet');
const toXMLString = require('./XMLString');

// inner workbook
function Workbook(outerWorkbook) {
	this.outerWorkbook = outerWorkbook;
	this.common = new Common();
	this.styles = this.common.styles;
	this.images = this.common.images;

	this.relations = new Relations(this.common);
	this.relations.add(this.styles, 'stylesheet');
}

Workbook.prototype = {
	addWorksheet(config) {
		config = _.defaults(config, {
			name: this.common.getNewWorksheetDefaultName()
		});
		const {outerWorksheet, worksheet} = createWorksheet(this.outerWorkbook, this.common, config);
		this.common.addWorksheet(worksheet);
		this.relations.add(worksheet, 'worksheet');

		return outerWorksheet;
	},

	/**
	 * Turns a workbook into a downloadable file.
	 * options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
	 */
	save(options) {
		const zip = new JSZip();
		const canStream = !!Readable;

		this.generateFiles(zip, canStream);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},
	saveAsNodeStream(options) {
		const zip = new JSZip();
		const canStream = !!Readable;

		this.generateFiles(zip, canStream);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	},

	generateFiles(zip, canStream) {
		this.prepareWorksheets();

		this.saveWorksheets(zip, canStream);
		this.saveTables(zip);
		this.saveImages(zip);
		this.saveDrawings(zip);
		this.saveStyles(zip);
		this.saveStrings(zip, canStream);
		zip.file('[Content_Types].xml', this.createContentTypes());
		zip.file('_rels/.rels', this.createWorkbookRelationship());
		zip.file('xl/workbook.xml', this.saveWorkbook());
		zip.file('xl/_rels/workbook.xml.rels', this.relations.save());
	},
	saveWorkbook() {
		return toXMLString({
			name: 'workbook',
			ns: 'spreadsheetml',
			attributes: [
				['xmlns:r', util.schemas.relationships]
			],
			children: [
				this.bookViewsXML(),
				this.sheetsXML(),
				this.definedNamesXML()
			]
		});
	},
	bookViewsXML() {
		const activeTab = this.common.getActiveWorksheetIndex();

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
	},
	sheetsXML() {
		const maxWorksheetNameLength = 31;
		const children = this.common.worksheets.map((worksheet, index) => {
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
					['r:id', this.relations.getId(worksheet)],
					['state', worksheet.getState()]
				]
			});
		});

		return toXMLString({
			name: 'sheets',
			children
		});
	},
	definedNamesXML() {
		const isPrintTitles = this.common.worksheets.some(
			worksheet => worksheet.printTitles && (worksheet.printTitles.topTo >= 0 || worksheet.printTitles.leftTo >= 0)
		);

		if (isPrintTitles) {
			const children = [];

			this.common.worksheets.forEach((worksheet, index) => {
				const entry = worksheet.printTitles;

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
	},
	prepareWorksheets() {
		this.common.worksheets.forEach(worksheet => {
			worksheet.prepare();
		});
	},
	saveWorksheets(zip, canStream) {
		this.common.worksheets.forEach(worksheet => {
			zip.file(worksheet.path, worksheet.save(canStream));
			zip.file(worksheet.relationsPath, worksheet.relations.save());
		});
	},
	saveTables(zip) {
		this.common.tables.forEach(table => {
			zip.file(table.path, table.save());
		});
	},
	saveImages(zip) {
		_.forEach(this.images.getImages(), image => {
			zip.file(image.path, image.data, {base64: true, binary: true});
			image.data = null;
		});
		this.images.removeImages();
	},
	saveDrawings(zip) {
		this.common.drawings.forEach(drawing => {
			zip.file(drawing.path, drawing.save());
			zip.file(drawing.relationsPath, drawing.relations.save());
		});
	},
	saveStyles(zip) {
		zip.file('xl/styles.xml', this.styles.save());
	},
	saveStrings(zip, canStream) {
		if (this.common.strings.isStrings()) {
			this.relations.add(this.common.strings, 'sharedStrings');
			zip.file('xl/sharedStrings.xml', this.common.strings.save(canStream));
		}
	},
	createContentTypes() {
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
		if (this.common.strings.isStrings()) {
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

		this.common.worksheets.forEach((worksheet, index) => {
			children.push(toXMLString({
				name: 'Override',
				attributes: [
					['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'],
					['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
				]
			}));
		});
		this.common.tables.forEach((table, index) => {
			children.push(toXMLString({
				name: 'Override',
				attributes: [
					['PartName', '/xl/tables/table' + (index + 1) + '.xml'],
					['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']
				]
			}));
		});
		_.forEach(this.common.images.getExtensions(), (contentType, extension) => {
			children.push(toXMLString({
				name: 'Default',
				attributes: [
					['Extension', extension],
					['ContentType', contentType]
				]
			}));
		});
		this.common.drawings.forEach((drawing, index) => {
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
	},
	createWorkbookRelationship() {
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
};

module.exports = Workbook;
