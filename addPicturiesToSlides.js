function addPicturiesToSlides() {
  var presentation = SlidesApp.getActivePresentation();
    console.log(presentation.getName()); // Starting script

  var maxWidth = 70; // Example: 50% of slide width
  var maxHeight = 100; // Example: 50% of slide height
  var slides = presentation.getSlides();
  var folderId = '1D_12WrwhX8p9PDwKggBOAPmH9v8FUtC9'; // Replace with your actual Google Drive folder ID
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var folderId2 = '18JFnTh89XY3g-3OH1ynexOTrSa9OX9au'; // Replace with your actual Google Drive folder ID
  var folder2 = DriveApp.getFolderById(folderId2);
  var files2 = folder2.getFiles();

  console.log(2); // Got slides and files

  slides.forEach(function(slide, index) {
    console.log("Processing slide #" + (index + 1));
    var rundomizer = index % 2 === 0 ? 2 : 1;;

    var shapes = slide.getShapes();
    shapes.forEach(function(shape) {
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        var text = shape.getText().asString();
                  console.log(text); 
        // Check if the shape's text matches the placeholder text
        if (text.includes('pic') ) {
        
        if (text.includes('{pic}') && files.hasNext()) {
          console.log(3); // Found placeholder and have file to insert
          var file = files.next();
          addPicToSlide(file, slide, shape, maxWidth, maxHeight);
          console.log(4); // Image inserted
        }
        if (text.includes(`{pic${rundomizer}}`) && files.hasNext()) {
          console.log(3); // Found placeholder and have file to insert
          var file = files.next();
          addPicToSlide(file, slide, shape, maxWidth, maxHeight);
          console.log(4); // Image inserted
        }


          if (text.includes(`{smallpic${rundomizer}}`) && files2.hasNext()) {
          console.log(3); // Found placeholder and have file to insert
          var file = files2.next();
          addPicToSlide(file, slide, shape, 35, 35);
          console.log(4); // Image inserted
        }

      shape.remove(); // Remove the placeholder text box

        }


      
      }
    });
  });

  // Helper function to add a picture to a slide
  function addPicToSlide(file, slide, shape, maxWidth, maxHeight) {
    var imageUrl = "https://drive.google.com/uc?export=view&id=" + file.getId();
    console.log("Inserting image from URL: " + imageUrl);
    var image = slide.insertImage(imageUrl);
    
    console.log(file.getName()); // Image added, now adjusting size
    
    // Initial dimensions
    var imageWidth = image.getWidth();
    var imageHeight = image.getHeight();

    // Calculate the scaling factor to fit the image within maxWidth and maxHeight, maintaining aspect ratio
    var scaleWidth = maxWidth / imageWidth;
    var scaleHeight = maxHeight / imageHeight;
    var scale = Math.min(scaleWidth, scaleHeight);

    // Apply scale; checks if scaling is necessary
    if (scale < 1) {
      imageWidth *= scale;
      imageHeight *= scale;
      image.setWidth(imageWidth);
      image.setHeight(imageHeight);
    }

    // Center the image over the placeholder
    var centerX = shape.getLeft() + (shape.getWidth() - imageWidth) / 2;
    var centerY = shape.getTop() + (shape.getHeight() - imageHeight) / 2;
    image.setLeft(centerX);
    image.setTop(centerY);

    console.log(7); // Image positioned and sized
  }

  console.log(8); // Script complete
}
