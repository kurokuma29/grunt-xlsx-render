module.exports = function(grunt) {
	grunt.initConfig({
		xlsx_render: {
			digit: 5,
			data: [
				{
					input: 'sample.xlsx',
					sheet: 'Sheet1',
					template: 'template.html',
					dest: 'output/test1/{{Regions}}_{{Prefectures}}.html',
				},
				{
					input: 'sample.xlsx',
					sheet: 'Sheet1',
					template: 'template.html',
					dest: 'output/test2/',
				},
				{
					input: 'sample.xlsx',
					sheet: 'Sheet1',
					setFileData: function(xlsxData){
						var fileData = [], pageData;
						for (var idx in xlsxData){
							pageData = xlsxData[idx];
							fileData.push({
								data: pageData,
								template: 'template.html',
								dest: 'output/test3/' + pageData.PrefecturalCapital + '.html'
							});
						}
						return fileData;
					}
				}
			]
		}
	});
	grunt.task.loadTasks('tasks');
};