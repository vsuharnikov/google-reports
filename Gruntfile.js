module.exports = function (grunt) {
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        //
        'gs-import': {
            options: {
                destDir: 'src',
                importFilePath: 'dist/graffity.json'
            }
        },
        'gs-export': {
            options: {
                srcDir: 'src',
                exportDir: 'dist',
                filePrefix: 'graffity'
            }
        }
    });
};