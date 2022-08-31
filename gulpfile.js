var gulp = require("gulp"),
merge = require("merge-stream"),
rimraf = require("rimraf");

var paths = {
    webroot: "./wwwroot/",
    node_modules: "./node_modules/"
};

paths.libDest = paths.webroot + "lib/";

gulp.task("libs", function () {
    var react = gulp.src(paths.node_modules + "react/umd/react.production.min.js")
        .pipe(gulp.dest(paths.libDest + "react"));
    var reactdom = gulp.src(paths.node_modules + "react-dom/umd/react-dom.production.min.js")
        .pipe(gulp.dest(paths.libDest + "react-dom"));

    return merge(react, reactdom);
});