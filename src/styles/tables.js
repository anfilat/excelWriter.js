'use strict';

const _ = require('lodash');
const StylePart = require('./stylePart');
const toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
const ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn',
	'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
	'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
const SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx
function Tables(styles) {
	StylePart.call(this, styles, 'tableStyles', 'table');

	this.saveEmpty = false;
}

Tables.prototype = _.merge({}, StylePart.prototype, {
	canon(format) {
		const result = {};
		const styles = this.styles;

		_.forEach(format, (value, key) => {
			if (_.includes(ELEMENTS, key)) {
				let style;
				let size = null;

				if (value.style) {
					style = styles.tableElements.add(value.style);
					if (value.size > 1 && _.includes(SIZED_ELEMENTS, key)) {
						size = value.size;
					}
				} else {
					style = styles.tableElements.add(value);
				}
				result[key] = {
					style,
					size
				};
			}
		});

		return result;
	},
	saveCollectionExt(attributes) {
		if (this.styles.defaultTableStyle) {
			attributes.push(['defaultTableStyle', this.styles.defaultTableStyle]);
		}
	},
	//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx
	saveFormat(format, styleFormat) {
		const styles = this.styles;
		const attributes = [
			['name', styleFormat.name],
			['pivot', 0]
		];
		const children = [];

		_.forEach(format, (value, key) => {
			const style = value.style;
			const size = value.size;
			const attributes = [
				['type', key],
				['dxfId', styles.tableElements.getId(style)]
			];

			if (size) {
				attributes.push(['size', size]);
			}

			children.push(toXMLString({
				name: 'tableStyleElement',
				attributes
			}));
		});
		attributes.push(['count', children.length]);

		return toXMLString({
			name: 'tableStyle',
			attributes,
			children
		});
	}
});

module.exports = Tables;
