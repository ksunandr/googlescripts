function addPicturesToSlides() {
  var presentation = SlidesApp.getActivePresentation();
  console.log(presentation.getName()); // Log the start of the script

  var slides = presentation.getSlides();

// Simplify folder and file retrieval
var folders = {
  big: DriveApp.getFolderById('1D_12WrwhX8p9PDwKggBOAPmH9v8FUtC9'), // main folder
  small: DriveApp.getFolderById('18JFnTh89XY3g-3OH1ynexOTrSa9OX9au'), // small images
  middle: DriveApp.getFolderById('1VGTR9908WUTcS9tCaCLU_tWiJfN1hO1A') // middle images
};

var fileArrays = {
  big: collectFiles(folders['big'].getFiles()), // Collect and then shuffle
  small: collectFiles(folders['small'].getFiles()),
  middle: collectFiles(folders['middle'].getFiles())
};

// Shuffle each array of files
shuffleArray(fileArrays.big);
shuffleArray(fileArrays.small);
shuffleArray(fileArrays.middle);
var i=0;

  slides.forEach(function(slide, index) {
      if (i < fileArrays.big.length){
    console.log("Processing slide #" + (index + 1));
    var randomizer = index % 2 === 0 ? 1 : 2;

                
    slide.getShapes().forEach(function(shape) {
    

      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        var text = shape.getText().asString();
       

        if (text.includes('pic')) {
           console.log("text " + (text));
          var fileType = '';
          if (text.includes('{pic}')) fileType = 'big';
          else if (text.includes(`{pic${randomizer}}`)) fileType = 'big';
          else if (text.includes(`{midpic${randomizer}}`)) fileType = 'middle';
          else if (text.includes(`{smallpic${randomizer}}`)) fileType = 'small';
          else if (text.includes(`{smallpic}`)) fileType = 'small';
         console.log("fileType " + (fileType));

          var file;

          if (fileType ) {           

          if (fileType =='middle' && i >=fileArrays.middle.length ) {
              file = fileArrays['middle'][i%fileArrays.middle.length];                
          } else {
            if (fileType =='small' && i >=fileArrays.small.length) {
              file = fileArrays['small'][i%fileArrays.small.length];
            } 
            if (fileType =='big' && i >=fileArrays.big.length) {
              file = fileArrays['big'][i%fileArrays.big.length];
            } else {
              file = fileArrays[fileType][i];
            }          
          }
            
            console.log("file " + (file));        

            
            var size;
            if (fileType === 'small') {
              size = { width: 30, height: 35 };
            } else if (fileType === 'middle') {
              size = { width: 60, height: 75 };
            } else {
              size = { width: 100, height: 140 };
            }


            addPicToSlide(file, slide, shape, size.width, size.height);

            
           
                        i++;

          }
           shape.remove(); // Remove the placeholder text box
           
        }
      }
                
    });
  }
  });

  // console.log("end " +countItemsInIterator(iterators['big'])); // Indicate script completion
}


function collectFiles(iterator) {
  var files = [];
  while (iterator.hasNext()) {
    files.push(iterator.next());
  }
  return files;
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    let j = Math.floor(Math.random() * (i + 1)); // random index from 0 to i
    [array[i], array[j]] = [array[j], array[i]]; // swap elements
  }
}


function countItemsInIterator(iterator) {
  var count = 0;
  while (iterator.hasNext()) {
    iterator.next();
    count++;
  }
  return count;
}

function addPicToSlide(file, slide, shape, maxWidth, maxHeight) {
  var imageUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
  var image = slide.insertImage(imageUrl);

  var imageWidth = image.getWidth();
  var imageHeight = image.getHeight();

  var scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight);

  if (scale < 1) {
    imageWidth *= scale;
    imageHeight *= scale;
    image.setWidth(imageWidth);
    image.setHeight(imageHeight);
  }

  var centerX = shape.getLeft() + (shape.getWidth() - imageWidth) / 2;
  var centerY = shape.getTop() + (shape.getHeight() - imageHeight) / 2;
  image.setLeft(centerX);
  image.setTop(centerY);

}
