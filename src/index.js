'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const JSZip = require('jszip');
const Workbook = require('./workbook');
const createWorksheet = require('./worksheet');

// Excel workbook API. Outer workbook
function createWorkbook() {
	const workbook = new Workbook();
	const common = workbook.common;
	const styles = common.styles;
	const images = common.images;
	const relations = workbook.relations;

	return {
		addWorksheet(config) {
			config = _.defaults(config, {
				name: common.getNewWorksheetDefaultName()
			});
			const {outerWorksheet, worksheet} = createWorksheet(this, common, config);
			common.addWorksheet(worksheet);
			relations.addRelation(worksheet, 'worksheet');

			return outerWorksheet;
		},

		addFormat(format, name) {
			return styles.addFormat(format, name);
		},
		addFontFormat(format, name) {
			return styles.addFontFormat(format, name);
		},
		addBorderFormat(format, name) {
			return styles.addBorderFormat(format, name);
		},
		addPatternFormat(format, name) {
			return styles.addPatternFormat(format, name);
		},
		addGradientFormat(format, name) {
			return styles.addGradientFormat(format, name);
		},
		addNumberFormat(format, name) {
			return styles.addNumberFormat(format, name);
		},
		addTableFormat(format, name) {
			return styles.addTableFormat(format, name);
		},
		addTableElementFormat(format, name) {
			return styles.addTableElementFormat(format, name);
		},
		setDefaultTableStyle(name) {
			styles.setDefaultTableStyle(name);
			return this;
		},

		addImage(data, type, name) {
			return images.addImage(data, type, name);
		},

		/**
		 * Turns a workbook into a downloadable file.
		 * options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
		 */
		save(options) {
			const zip = new JSZip();
			const canStream = !!Readable;

			workbook.generateFiles(zip, canStream);
			return zip.generateAsync(_.defaults(options, {
				compression: 'DEFLATE',
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				type: 'base64'
			}));
		},

		saveAsNodeStream(options) {
			const zip = new JSZip();
			const canStream = !!Readable;

			workbook.generateFiles(zip, canStream);
			return zip.generateNodeStream(_.defaults(options, {
				compression: 'DEFLATE'
			}));
		}
	};
}

module.exports = {createWorkbook};
