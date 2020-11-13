'use strict';

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    $(function() {
      $('p').text('Hello World!!');
    });
  }
});
