'use strict';

const Workbook = require('./workbook');

// Excel workbook API
function createWorkbook() {
	const outerWorkbook = {
		addWorksheet(config) {
			return workbook.addWorksheet(config);
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
			return images.add(data, type, name);
		},

		save(options) {
			return workbook.save(options);
		},

		saveAsNodeStream(options) {
			return workbook.saveAsNodeStream(options);
		}
	};
	const workbook = new Workbook(outerWorkbook);
	const styles = workbook.styles;
	const images = workbook.images;

	return outerWorkbook;
}

module.exports = {createWorkbook};
