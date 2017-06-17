'use strict';

module.exports = function (excel) {
	const data = [
		{data: ['left 1', 'tempor nisi', 'aliqua tempor', 'enim officia', 'ipsum dolor'], height: 50},
		{data: ['left 2', 'qui Lorem', 'officia consectetur', 'pariatur mollit', 'ipsum reprehenderit'], height: 50},
		{data: ['left 3', 'ullamco proident', 'officia cupidatat', 'minim proident', 'mollit exercitation'], height: 50}
	];

	return excel.createWorkbook()
		.addWorksheet()
		.setPageMargin({
			left: 0.8,
			right: 0.8,
			top: 1.75,
			bottom: 1.75,
			header: 0.7,
			footer: 0.9
		})
		.setPageOrientation('landscape')
		.setHeader(['', 'Top center text', 'Page &P of &N'])
		.setFooter([
			{text: 'left', font: 'Arial', bold: true, italic: true, underline: true, fontSize: 15},
			['before ', {text: 'center text', bold: true}, ' after'],
			'&F'
		])
		.setPrintTitleTop(2)
		.setPrintTitleLeft(1)
		.setColumns([
			{width: 15},
			{width: 30},
			{width: 30},
			{width: 30},
			{width: 30}
		])
		.setData([
			[1, 2, 3, 4, 5],
			['text']
		])
		.setData(data)
		.setData(data)
		.setData(data)
		.setData(data)
		.end();
};
