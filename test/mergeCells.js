'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet()
		.setData([
			[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 260, 'Dolore anim'],
			[2541, {value: 'Nullam aliquet mi et nunc tempus rutrum.', colspan: 2, rowspan: 3}, 'Dolore anim'],
			[2541, 'Dolore anim']
		])
		.setColumns([
			{type: 'number', width: 10},
			{type: 'string', width: 60},
			{type: 'number'},
			{type: 'string', width: 20}
		])
		.end();
};
