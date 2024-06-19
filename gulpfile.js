
const gulp = require('gulp');
const nodemon =require('gulp-nodemon');
const clean        = require('gulp-clean');
const pump         = require('pump');
const babel = require('gulp-babel');
const spawn = require('child_process').spawn;

let jsScript ='node';
if(process.env.npm_config_argv !== undefined && process.env.npm_config_argv.indexOf('debug') !== -1){
    jsScript = 'node debug';
}

gulp.task('nodemon',function(){
    return nodemon({
        script:'build/server.js',
        execMap:{
            js:jsScript
        },
        verbose:true,
        ignore:['dist/*.js','node_modules/**/node_modules','gulpfile.js'],
        env:{
          NODE_ENV:'development'
        },
        ext:'js json'
    })
});

gulp.task('default',gulp.series('nodemon',function(){
    console.log('熱加載已經啟動...')
}))

gulp.task('babel', () => {
    return gulp.src('./src/**')
        .pipe(babel({
            "presets": [["env",{
                "targets":{
                    "node":"8.11.3"
                }
            }], "stage-2"],
            "plugins": ["transform-runtime"]
        }))
        .pipe(gulp.dest('./dist'));
});

gulp.task('clean',function(cb){
    pump([
        gulp.src('./dist'),
        clean()
    ], cb)
});


gulp.task('build', gulp.series('clean','babel'));


