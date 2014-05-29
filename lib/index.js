// index.js
var fs = require('fs');
var path = require('path');
var officegen = require('officegen');

// PowerPoint
var pptx = officegen({
    'type': 'pptx',
    'onend': function(written) {
        console.log('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n');
    },
    'onerr': function(err) {
        console.log(err);
    }
});
var reg = /.*¥.gif|png|jpg|jpeg|bmp$/i;

// generate pptx file from images.
function generate(dir, fileName, callback){
    var existCallback = (typeof callback == "function");

    try{
        fs.readdir(dir, function(err, files){
            if(err){
                if(existCallback){
                    callback(err);
                }
            }

            var images = [];
            files.forEach(function(file){
                // filter image file
                if(reg.test(file)){
                    images.push(path.join(dir, file));
                }
            });
            makeSlide(fileName, images, callback);
        });
    }catch(e){
        if(existCallback){
            callback(e);
        }
    }
}

// add slides and output pptx file.
function makeSlide(file, images, callback){
    var existCallback = (typeof callback == "function");

    // add slides
    images.forEach(function(image){
        if(image.length > 0){
            console.log(image);
            var slide = pptx.makeNewSlide();
            slide.addImage(image, { y: 'c', x: 'c', cy: '70%', cx: '70%' });
        }
    });

    // output PPTX
    var out = fs.createWriteStream(file);

    pptx.generate(out, {
        'finalize': function(written) {
            console.log ('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );

            if(existCallback){
                callback(null, file);
            }
        },
        'error': function(err) {
            if(existCallback){
                callback(err);
            }
        }
    });
}

module.exports = {
    generate: generate
};