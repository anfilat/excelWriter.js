'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setData([
			['First', 260],
			['Second', 42],
			['Total', {formula: 'B1+B2'}]
		])
		.end();
};
