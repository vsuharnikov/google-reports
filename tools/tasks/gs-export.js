/*
 * Copyright (c) 2014 Vyatcheslav Suharnikov, contributors
 * Licensed under the MIT license.
 */

'use strict';

// DOESN'T WORK, BECAUSE OF GOOGLE!

var path = require('path');

module.exports = function (grunt) {
    /**
     * Paths should be relative to the Gruntfile.js.
     *
     * @options options {string} srcDir                 A path to the directory, files imported in.
     * @options options {string} exportDir   A path to the file, exported from Google Drive.
     *
     * @see https://developers.google.com/apps-script/import-export
     */
    grunt.registerTask('gs-export', 'Export a google script project', function () {
        var options = this.options();

        if (!options.srcDir) {
            return grunt.log.error('Please specify the "srcDir" in options.');
        }

        var metaFilePath = path.join(options.srcDir, '.meta.json');
        if (!grunt.file.exists(metaFilePath)) {
            return grunt.log.error('The meta file doesn\'n exist.');
        }

        var srcFiles = grunt.file.expand(path.join(options.srcDir, '*.js'));
        if (srcFiles.length == 0) {
            return grunt.log.error('There are no files to export.');
        }

        var meta = grunt.file.readJSON(metaFilePath),
            result = {files: []};

        // Collect files content.
        var allFilesHaveMeta = srcFiles.every(function (filePath) {
            var fileMeta = meta[filePath];
            if (!fileMeta) {
                grunt.log.error('The file "' + filePath + '" hasn\'t meta information.');
                return false;
            }

            fileMeta.source = grunt.file.read(filePath, {encoding: 'utf8'});
            result.files.push(fileMeta);
            return true;
        });

        if (!allFilesHaveMeta) {
            return grunt.log.error('There are files without meta. Exporting is failed.');
        }

        var exportedFilePath = path.join(options.exportDir, getExportFileName(options.filePrefix));
        grunt.file.write(
            exportedFilePath,
            JSON.stringify(result, null, '\t'),
            {encoding: 'utf8'}
        );

        grunt.log.ok('The result is written to the "' + exportedFilePath + '" file.');
    });

    function getExportFileName(prefix) {
        return prefix + '_' + new Date().toJSON().replace(/[\:\.]/g, '') + '.json';
    }
};
