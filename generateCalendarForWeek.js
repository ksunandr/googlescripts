function generateCalendarForWeek() { // https://docs.google.com/presentation/d/1Qkz1gXCBa8f7z56N0OJIwG3lvSt3feFaXThAEWPKagI/edit#slide=id.g2b8a3099163_0_317
  var presentation = SlidesApp.getActivePresentation();
  var templateSlide = presentation.getSlides()[0];
  var startDay = 13;
  var startMonth = 4; //may
  var year = 2024;
  
  var startDate = new Date(year, startMonth, startDay);
  var totalDays = 500;
  var currentDay = startDate;

  for (var i = 0; i < totalDays; i += 7) {
    var newSlide = presentation.appendSlide(templateSlide);


    for (var weekDay = 1; weekDay < 8; weekDay++) {

       var dateStr = currentDay.getDate().toString();
       var monthNumberStr = (currentDay.getMonth() + 1).toString();

       newSlide.replaceAllText(`{{day${weekDay}}}`, dateStr);
       newSlide.replaceAllText(`{{month${weekDay}}}`, monthNumberStr);


       currentDay.setDate(currentDay.getDate() + 1);

  }
  }
     
 

}
