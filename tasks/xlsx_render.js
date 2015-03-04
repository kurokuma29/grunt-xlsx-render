module.exports = function(grunt) {
	var Xlsx = require('xlsx'),
	Mustache = require('mustache'),
	Promise = require('es6-promise').Promise,
	Util = require('util'),

	DEFAULT_OPTIONS = {
		digit: 5,
		default_extension: '.html'
	};

	function xlsxRender(){
		this.options = grunt.util._.extend(DEFAULT_OPTIONS, grunt.config('xlsx_render'));
	}

	xlsxRender.prototype._chkPara = function(data){
		var toString = Object.prototype.toString;
		var arMsg = [];

		// input file
		if(toString.call(data.input) !== '[object String]' || data.input === ''){
			arMsg[arMsg.length] = '* Input file is not set.';
		}else{
			// input file check
			if(!grunt.file.exists(data.input)){
				arMsg[arMsg.length] = '* Input file does not exist.';
			}
		}
		// sheet
		if(toString.call(data.sheet) !== '[object String]' || data.sheet === ''){
			arMsg[arMsg.length] = '* Sheet name has not been set.';
		}
		// setFileData
		if(toString.call(data.setFileData) !== '[object Function]'){
			// To check if it is not a function
			// template
			if(toString.call(data.template) !== '[object String]' || data.template === ''){
				arMsg[arMsg.length] = '* Template file has not been set.';
			}
			// destination
			if(toString.call(data.dest) !== '[object String]' || data.dest === ''){
				arMsg[arMsg.length] = '* Output destination is not set.';
			}
		}
		return arMsg;
	};

	xlsxRender.prototype._showError = function(arErr){
		grunt.log.error('Parameter is incorrect.');
		Object.keys(arErr).forEach(function(idx){
			grunt.log.writeln('Parameter index ' +  idx + ' -------------------------------');
			for(var i = 0, j = arErr[idx].length; i < j; i++){
				grunt.log.writeln(arErr[idx][i].red);
			}
		});
	};

	xlsxRender.prototype._getReplaceCharacters = function(string, pageData){
		var arDest, ret;
		ret = string;
		arDest = string.match(/\{\{(.*?)\}\}/g);
		if(Util.isArray(arDest)){
			for(var i=0, j=arDest.length; i<j; i++){
				ret = ret.split(arDest[i]).join(pageData[arDest[i].replace('{{', '').replace('}}', '')]);
			}
		}
		return ret;
	};

	xlsxRender.prototype._padLeft = function(number, digit){
		var pad = '';
		for(var i=0; i<digit; i++){
			pad += '0';
		}
		return (pad + (number)).slice(digit * -1);
	};

	xlsxRender.prototype._getXlsx = function(input, sheet){
		// read data
		var sheetData = Xlsx.readFile(input).Sheets[sheet];
		// to json
		return Xlsx.utils.sheet_to_json(sheetData);
	};

	xlsxRender.prototype._getFileData = function(param){
		return new Promise(function(resolve, reject){
			var toString = Object.prototype.toString;
			var fileData, jsonData;
			jsonData = this._getXlsx(param.input, param.sheet);
			if(toString.call(param.setFileData) === '[object Function]'){
				try{
					fileData = param.setFileData(jsonData);
					if(!Util.isArray(fileData)){
						reject(new Error('Return value of [setFileData] is invalid.'));
					}
					resolve(fileData);
				}catch(e){
					reject(new Error('[setFileData] is invalid -> ' + e));
				}
			}else{
				var pageData, dest, template;
				var cnt = 0;
				var arDest = [];
				fileData = [];
				for (var idx in jsonData){
					pageData = jsonData[idx];
					cnt++;
					dest = param.dest;
					if(grunt.file.isMatch({ matchBase: true }, '*.*', dest)){
						// replace charactors
						dest = this._getReplaceCharacters(dest, pageData);
					}else{
						// auto increment
						if(dest.slice(-1) !== '/'){
							dest += '/';
						}
						dest += this._padLeft(cnt, this.options.digit) + this.options.default_extension;
					}
					template = this._getReplaceCharacters(param.template, pageData);
					// fileData
					fileData[fileData.length] = {
						data: pageData,
						template: template,
						dest: dest
					};
				}
				resolve(fileData);
			}
		}.bind(this));
	};

	xlsxRender.prototype._doRender = function(data){
		return new Promise(function(resolve, reject){
			try{
				var dataIdx = 1;
				var cnt = 0;
				data.map(function(grpData){
					grunt.log.writeln('Dataset Index: '.yellow + String(dataIdx).yellow);
					grpData.map(function(fileData){
						var templateData = '';
						// get template html
						if(!grunt.file.exists(fileData.template)){
							reject(new Error('Template File does not exist.'));
						}
						templateData = grunt.file.read(fileData.template);
						grunt.log.writeln('Output file: ' + fileData.dest);
						grunt.file.write(fileData.dest, Mustache.to_html(templateData, fileData.data));
						grunt.log.ok('object: ' + String(Object.keys(fileData.data).length).cyan + '-key' + ', template: ' + fileData.template.cyan);
						cnt++;
					});
					grunt.log.writeln();
					dataIdx++;
				});
				resolve(cnt);
			}catch(e){
				reject(e);
			}
		});
	};

	grunt.registerTask('xlsx_render', 'Generate files from Xlsx file.', function() {
		var renderer = new xlsxRender();
		var toString = Object.prototype.toString;
		var arErr = {};
		var arTmp;

		// Parameters check
		if(toString.call(renderer.options.data) === '[object Undefined]' || !Util.isArray(renderer.options.data) || renderer.options.data.length === 0){
			throw new Error('Parameters of the [data] is invalid.');
		}
		renderer.options.data.map(function(param, idx){
			arTmp = renderer._chkPara(param);
			if(arTmp.length > 0){
				arErr[idx + 1] = arTmp;
			}
		});
		if(Object.keys(arErr).length > 0){
			renderer._showError(arErr);
			throw new Error('Parameters of the [data] is invalid.');
		}

		// grunt done
		var done = this.async();

		// Output files
		Promise.all(renderer.options.data.map(function(param){
			return renderer._getFileData(param);
		})).then(function(res){
			return renderer._doRender(res);
		}).then(function(cnt){
			grunt.log.ok('Successfully all output files: ' + cnt);
			done();
		}).catch(function(e){
			if(e){
				grunt.log.error(e.toString());
				if (typeof e.stack === 'string') {
					e.stack.split('\n').filter(Boolean).slice(1).forEach(function(line) {
						grunt.verbose.error(line);
					});
				}
			}
			done(false);
		});
	});
};