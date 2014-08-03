/*
 * excel
 * https://github.com/leo/node-excel
 *
 * Copyright (c) 2014 leoxie
 * Licensed under the MIT license.
 */

'use strict';

var xlsx = require('node-xlsx'),
	nodeExcel = require('excel-export-fast'),
	moment = require('moment'),
	_ = require('./utility'),
	path = require('path');

function _v(row, index, format, valid) {
	var r = row[index];
	if (r) {
		if (valid && !valid(r)) {
			return {
				error: true,
				info: ' 列:' + index + ' 值:' + r.value
			};
		}
		if (format) return format(r);
		if (r.formatCode == '@') {
			if (_.isString(r.value)) {
				return r.value.trim();
			}
			return r.value;
		} else {
			return new moment(r.value).format('YYYY-MM-DD');
		}
	}
	return null;
}

function _n(data, name, v) {
	var names = name.split('.');
	var d = v;
	for (var i = 0, j = names.length - 1; i <= j; j--) {
		var n = names[j];
		var t = d;
		d = {};
		d[n] = t;
	}
	data = _.merge(data, d);
}

/**
 * 解析xlsx
 * @param  {[type]} file_name [description]
 * @param  {[type]} begin     [description]
 * @param  {[type]} format    格式为数组, 内容为{ name: 'xx', format: function(r){}, valid: function(r){}} 或 字符串
 * @return {[type]}           [description]
 */
var parse = function(file_name, begin, format) {
	var obj = xlsx.parse(file_name);
	var data = obj.worksheets[0].data;
	var result = [];
	var err = [];
	for (var i = begin; i < data.length; i++) {
		var row = data[i];
		if (row) {
			var d = {};
			for (var j = 0; j < format.length; j++) {
				var f = format[j];
				if (f == null) {
					continue;
				} else if (_.isObject(f)) {
					var _r = _v(row, j, f['format'], f['valid']);
					if (_r instanceof Object && _r.error) {
						_r['info'] = '行:' + i + _r['info'];
						err.push(_r);
						_n(d, f.name, null);
					} else {
						_n(d, f.name, _r);
					}
				} else {
					_n(d, f, _v(row, j));
				}
			}
			result.push(d);
		}
	}
	return {
		err: err.length == 0 ? null : err.map(function(r) {
			return r.info;
		}).join(' , '),
		result: result
	};
};

/**
 * 导出xlsx
 * @param  {[type]}   data   [description]
 * @param  {[type]}   format 格式为数组, 内容为{ name: 'xx', format: function(r){}, caption: 'xx'}
 * @param  {Function} cb     [description]
 * @return {[type]}          [description]
 */
var _export = function(data, format, cb) {
	var cols = [],
		rows = [];
	for (var i = 0; i < format.length; i++) {
		cols.push({
			caption: format[i].caption,
			type: 'string'
		});
	};
	for (var i = 0; i < data.length; i++) {
		var d = data[i];
		var _d = [];
		for (var j = 0; j < format.length; j++) {
			var f = format[j];
			var old_data = _.getPath(d, f.name);
			old_data = f.format ? f.format(old_data) : ('' + old_data);
			_d.push(old_data);
		};
		rows.push(_d);
	};
	var conf = {
		stylesXmlFile: path.join(__dirname, "styles.xml"),
		cols: cols,
		rows: rows
	};
	nodeExcel.execute(conf, cb);
};

exports.parse_xlsx = parse;
exports.export_xlsx = _export;

// _export([{
// 	a: '112312312312312312312312312312312312231312312312312312312',
// 	b: 'sdf',
// 	c: 'ww',
// 	z: {
// 		a: 'za',
// 		b: 'zb'
// 	}
// }, {
// 	a: 2,
// 	b: 'sdf',
// 	c: 'ww',
// 	z: {
// 		a: 'za',
// 		b: 'zb'
// 	}
// }], [{
// 	name: 'b',
// 	caption: 'sb',
// 	format: function(v) {
// 		return moment().format('YYYY-MM-DD');
// 	}
// }, {
// 	name: 'z.a',
// 	caption: 'sa'
// }, {
// 	name: 'z.b',
// 	caption: 'sc'
// }], function(err, r) {
// 	console.info(err);
// 	console.info(r);
// 	require('fs').writeFileSync(path.join(__dirname, '../test/test1.xlsx'), r);
// });



// require('fs').writeFileSync(path.join(__dirname, '../test/test.json'), JSON.stringify(
// 	parse(path.join(__dirname, '../test/test.xlsx'), 1, [{
// 			name: 'no',
// 			format: function(v) {
// 				return v.value + '1';
// 			},
// 			valid: function(v) {
// 				return v.value > 17283;
// 			}
// 		},
// 		'w.qudao',
// 		'z.dj_date',
// 		'w.wangdian',
// 		'z.zhongtuo',
// 		'z.w.tuozhang',
// 		'fengxian',
// 		'ruzhang',
// 		'shanghu_id',
// 		'd.shanghu_name',
// 		'd.a.shanghu_address',
// 		'zhuangji_address',
// 		'jingying',
// 		'shangquan',
// 		'shanghu_fuzeren',
// 		'feilv',
// 		'zhongduan',
// 		'zhuangji_num',
// 		'lianxiren',
// 		'phone',
// 		'sp_date',
// 		'jishenhao',
// 		'shibiema',
// 		'pos_phone'
// 	])
// ));