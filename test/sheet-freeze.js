'use strict';

module.exports = function (excel) {
	return excel.createWorkbook()
		.addWorksheet({
			freeze: {col: 1, row: 2}
		})
		.setData([
			[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 260, 1342372604000, 1342977404000],
			[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 260, 1342372604000, 1342977404000],
			[2541, 'Nullam aliquet mi et nunc tempus rutrum.', 260, 1342372604000, 1342977404000]
		])
		.end();
};
