# grunt-xlsx-render

> xlsx to html render

## Getting Started
This plugin requires Grunt `~0.4.5`

If you haven't used [Grunt](http://gruntjs.com/) before, be sure to check out the [Getting Started](http://gruntjs.com/getting-started) guide, as it explains how to create a [Gruntfile](http://gruntjs.com/sample-gruntfile) as well as install and use Grunt plugins. Once you're familiar with that process, you may install this plugin with this command:

```shell
npm install grunt-xlsx-render --save-dev
```

Once the plugin has been installed, it may be enabled inside your Gruntfile with this line of JavaScript:

```js
grunt.loadNpmTasks('grunt-xlsx-render');
```

## xlsx_render task

### Overview
In your project's Gruntfile, add a section named `xlsx_render` to the data object passed into `grunt.initConfig()`.

```js
grunt.initConfig({
  xlsx_render: {
    // Task options go here.
  }
});
```

### Options

#### options.digit
Type: `number`
Default value: `5`

Number of digits at the time of the output file name automatically generated.

#### options.default_extension
Type: `string`
Default value: `.html`

Extension of the time of the output file name automatically generated.

### Usage Examples

#### Preparation

##### Xlsx File
I prepare the xlsx file.
First row header.
Subsequent rows to create to be a value.
|no  |foo  |bar  |
|----|-----|-----|
|1   |Foo1 |Bar1 |
|2   |Foo2 |Bar2 |


##### Template File
How to write a template file is the same as the mustache.
```
<!DOCTYPE html>
<html>
  <head>
    <title>{{foo}}</title>
    <meta charset="utf-8">
  </head>
  <body>
    <h1>{{foo}}</h1>
    <div>{{bar}}</div>
  </body>
</html>
```

#### Example output 1
In this example, to use the value of xlsx the destination file name.

```
grunt.initConfig({
  xlsx_render: {
    data: [
      {
        input: 'sample.xlsx',
        sheet: 'Sheet1',
        template: 'template.html',
        dest: 'output/{{foo}}_{{bar}}.html'
      }
    ]
  }
});
```

#### Example output 2
In this example, to automatically generate a file name of the output destination.

```
grunt.initConfig({
  xlsx_render: {
    digit: 6,
    data: [
      {
        input: 'sample.xlsx',
        sheet: 'Sheet1',
        template: 'template.html',
        dest: 'output/'
      }
    ]
  }
});
```

#### Example output 3
In this example, I have created in my data set for output.
argument of setFileData is Json.

```
grunt.initConfig({
  xlsx_render: {
    data: [
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
              dest: 'output/' + pageData.foo + '.html'
            });
          }
          return fileData;
        }
      }
    ]
  }
});
```

#### Example output 4
Successively outputs can be.

```
grunt.initConfig({
  xlsx_render: {
    data: [
      {
        input: 'sample1.xlsx',
        sheet: 'Sheet1',
        template: 'template.html',
        dest: 'output/{{foo}}_{{bar}}.html'
      },
      {
        input: 'sample2.xlsx',
        sheet: 'Sheet1',
        template: 'template.html',
        dest: 'output/{{foo2}}_{{bar2}}.html'
      }
    ]
  }
});
```