/*
 * Copyright (c) 2014 Vyatcheslav Suharnikov, contributors
 * Licensed under the MIT license.
 */

'use strict';

var path = require('path');

module.exports = function (grunt) {

    /**
     * Paths should be relative to the Gruntfile.js.
     *
     * @options options {string} destDir        A path to the directory, files imported in.
     * @options options {string} importFilePath A path to the file, exported from Google Drive.
     *
     * @see https://developers.google.com/apps-script/import-export
     */
    grunt.registerTask('gs-import', 'Import a google script project', function () {
        var options = this.options();

        if (!options.destDir) {
            return grunt.log.error('Please specify the "destDir" in options.');
        }

        if (!options.importFilePath) {
            return grunt.log.error('Please specify the "importFilePath" in options.');
        }

        var archive = grunt.file.readJSON(options.importFilePath);
        if (!archive.files) {
            return grunt.log.error('An invalid archive: the "files" element is expected.');
        }

        var meta = {};

        // Import files to the destDir.
        archive.files.forEach(function (fileEntry) {
            var filePath = path.join(options.destDir, fileEntry.name + '.js'),
                wasWritten = grunt.file.write(filePath, fileEntry.source, {encoding: 'utf8'});

            meta[filePath] = {
                id: fileEntry.id,
                name: fileEntry.name,
                type: fileEntry.type
            };

            if (wasWritten) {
                grunt.log.ok('The file "' + filePath + '" was successfully imported.');
            } else {
                grunt.log.error('Can\'t import the file: "' + filePath + '".');
            }
        });

        // Write a meta information.
        var metaFilePath = path.join(options.destDir, '.meta.json');
        if (grunt.file.write(metaFilePath, JSON.stringify(meta, null, '\t'), {encoding: 'utf8'})) {
            grunt.log.ok('The meta-file was written.');
        } else {
            grunt.log.error('The meta-file wasn\'t written.');
        }
    });

};
