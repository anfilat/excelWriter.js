'use strict';

const Readable = require('stream').Readable;
const _ = require('lodash');
const JSZip = require('jszip');
const Workbook = require('./workbook');

module.exports = {
	createWorkbook: function () {
		return new Workbook();
	},

	/**
	 * Turns a workbook into a downloadable file.
	 * @param {Workbook} workbook - The workbook that is being converted
	 * @param {Object?} options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
	 */
	save: function (workbook, options) {
		const zip = new JSZip();
		const canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},

	saveAsNodeStream: function (workbook, options) {
		const zip = new JSZip();
		const canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	}
};
