module.exports = function(grunt) {

	grunt.initConfig({
		pkg: grunt.file.readJSON('package.json'),
		// Clean build directories
		clean: {
			src: ["public/css", "public/js"]
		},
		// JS linting
		jshint: {
			// 'build/js/lib/*.js',
			// Only JSHint custom source files, libraries can be hinted too if required
			files: ['gruntfile.js', 'build/js/src/*.js'],
			options: {
				// options here to override JSHint defaults
				globals: {
					jQuery: true,
					console: true,
					module: true,
					document: true
				}
			}
		},
		// JS concatenation
		concat: {
			options: {
			},
			dist: {
				src: ['build/js/lib/*.js', 'build/js/src/*.js'],
				dest: 'public/js/<%= pkg.name %>.js'
			}
		},
		// JS minification
		uglify: {
			options: {
				banner: '/*! <%= pkg.name %> <%= grunt.template.today("dd-mm-yyyy") %> */\n'
			},
			dist: {
				files: {
					'public/js/<%= pkg.name %>.min.js': ['<%= concat.dist.dest %>']
				}
			}
		},
		// Compass options
		compass: {
			dev: {
				options: {
					config: 'config.rb'
				}
			},
			dist: {
				options: {
					environment: 'production',
					config: 'config.rb'
				}
			}
		},
		// Optimise images

		// Watch command
		watch: {
			files: ['<%= jshint.files %>', 'build/scss/*.scss', 'build/scss/defaults/*.scss', 'build/scss/elements/*.scss', 'build/scss/partials/*.scss'],
			tasks: ['dev']
		}
	});

	// Load tasks
	grunt.loadNpmTasks('grunt-contrib-uglify');
	grunt.loadNpmTasks('grunt-contrib-jshint');
	grunt.loadNpmTasks('grunt-contrib-watch');
	grunt.loadNpmTasks('grunt-contrib-concat');
	grunt.loadNpmTasks('grunt-contrib-compass');
	grunt.loadNpmTasks('grunt-contrib-clean');

	// Set 2 basic tasks:
	// - dev for development (no cleaning of directories and uncompressed CSS and JS)
	// - default for production (cleans directories first and replaces with compressed CSS and JS)
	grunt.registerTask('dev', ['jshint', 'concat', 'compass:dev']);
	grunt.registerTask('default', ['clean', 'jshint', 'concat', 'uglify', 'compass:dist']);

};
