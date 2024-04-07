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

// Simplify file retrieval with iterators
var iterators = {
  big: folders['big'].getFiles(), // main images
  small: folders['small'].getFiles(), // small images
  middle: folders['middle'].getFiles() // middle images
};

  console.log("start " +countItemsInIterator(iterators['big'])); // Indicate script completion


  slides.forEach(function(slide, index) {
    console.log("Processing slide #" + (index + 1));
    var randomizer = index % 2 === 0 ? 1 : 2;

    slide.getShapes().forEach(function(shape) {
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        var text = shape.getText().asString();

        if (text.includes('pic')) {
          var fileType = '';
          if (text.includes('{pic}')) fileType = 'big';
          else if (text.includes(`{pic${randomizer}}`)) fileType = 'big';
          else if (text.includes(`{midpic${randomizer}}`)) fileType = 'middle';
          else if (text.includes(`{smallpic${randomizer}}`)) fileType = 'small';


          if (fileType =='middle' && !iterators['middle'].hasNext()) {
            iterators['middle'] = folders['middle'].getFiles();
          }
          if (fileType =='small' && !iterators['small'].hasNext()) {
            iterators['small'] = folders['small'].getFiles();
          }

          if (fileType && iterators['big'].hasNext()) {

            var file = iterators[fileType].next();            
            
            var size;
            if (fileType === 'small') {
              size = { width: 30, height: 35 };
            } else if (fileType === 'middle') {
              size = { width: 60, height: 75 };
            } else {
              size = { width: 90, height: 130 };
            }

            addPicToSlide(file, slide, shape, size.width, size.height);
          }
          
          shape.remove(); // Remove the placeholder text box
        }
      }
    });
  });

  console.log("end " +countItemsInIterator(iterators['big'])); // Indicate script completion
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
