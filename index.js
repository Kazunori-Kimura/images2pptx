// index.js
var fs = require('fs');
var path = require('path');
var officegen = require('officegen');
var br = String.fromCharCode(13) + String.fromCharCode(10);

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

// make Slides
var data = fs.readFileSync('list.txt', {encoding: 'UTF-8'});
var lines = data.split(br);

lines.forEach(function(line){
    if(line.length > 0){
        var slide = pptx.makeNewSlide();

        slide.addImage(line, { y: 'c', x: 'c', cy: '70%', cx: '70%' });
    }
});

// output PPTX
var out = fs.createWriteStream('out.pptx');

pptx.generate(out, {
    'finalize': function(written) {
        console.log ('Finish to create a PowerPoint file.\nTotal bytes created: ' + written + '\n' );
                        },
    'error': function(err) {
        console.log(err);
    }
});
