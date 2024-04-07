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

    // Simplify folder and file retrieval
  var iterators = {
   big: folders[big].getFiles(), // main folder
   small: folders[small].getFiles(), // small images
    middle:folders[middle].getFiles() // middle images
  };


  console.log(2); // Indicate that slides and files are ready for processing

  slides.forEach(function(slide, index) {
    console.log("Processing slide #" + (index + 1));
    var randomizer = index % 2 === 0 ? 2 : 1;

    slide.getShapes().forEach(function(shape) {
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        var text = shape.getText().asString();
        console.log(text); // Log the text found
        
        var fileType = '';
        if (text.includes('{pic}')) fileType = 'default';
        else if (text.includes(`{pic${randomizer}}`)) fileType = 'default';
        else if (text.includes(`{midpic${randomizer}}`)) fileType = 'middle';
        else if (text.includes(`{smallpic${randomizer}}`)) fileType = 'small';

        if (fileType && iterators[fileType].hasNext()) {
          var file = iterators[fileType].next();
          
          console.log("Processing file " + file.getName() +" " + fileType);
          
          var size;

          if (fileType === 'small') {
            size = { width: 25, height: 35 };
          } else if (fileType === 'middle') {
            size = { width: 45, height: 55 };
          } else {
            size = { width: 70, height: 100 };
          }


          addPicToSlide(file, slide, shape, size.width, size.height);
          shape.remove(); // Remove the placeholder text box
        }
      }
    });
  });

  console.log(8); // Indicate script completion
}

function addPicToSlide(file, slide, shape, maxWidth, maxHeight) {
  var imageUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
  console.log("Inserting image from URL: " + imageUrl);
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

  console.log(7); // Indicate that the image has been positioned and sized
}
