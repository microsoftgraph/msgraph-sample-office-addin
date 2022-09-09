// TEMPORARY CODE TO VERIFY ADD-IN LOADS
'use strict';

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(function() {
      $('p').text('Hello World!!');
    });
  }
});