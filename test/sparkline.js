'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.addSparkline({cell: 'A1', beginCell: 'C1', endCell: 'J1'})
		.addSparkline({cell: {c: 1, r: 3}, beginCell: {c: 3, r: 3}, endCell: {c: 10, r: 3}})
		.setData([
			[null, null, 2, 3, 44, -6, 7, 12, 25, -12],
			[],
			[{colspan: 2, rowspan: 2}, 5, 12, 17, 23, 34, 35, 39, 42]
		])
		.end();
};
