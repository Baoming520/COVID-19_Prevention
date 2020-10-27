$(() => {
  // The status checking page has been loaded successfully.
  getContent();
  setInterval(getContent, 3000);
});

const REQUEST_HTML = '<font style="font-weight:bold" color="blue">%TEXT%</font>';

var getStatusHtml = (status) => {
  const EXEC_IMAGE_HTML = '<img id="download" src="./images/check/loading.gif" height="80%" width="auto" />';
  const DONE_IMAGE_HTML = '<img id="download" src="./images/check/done.png" height="80%" width="auto" />';
  var html = '';
  if(status === 'Pending' || status === 'In Progress'){
    html += status + '&nbsp;' + EXEC_IMAGE_HTML;
  }
  else if(status === 'Completed'){
    html += status + '&nbsp;' + DONE_IMAGE_HTML;
  }
  else{
    html += status;
  }

  return html;
};

// ./images/check/loading.gif
var getContent = () => {
  $.get('/actions', (actions) => {
    $('#download_request').html(REQUEST_HTML.replace('%TEXT%', actions.downloadRequest));
    $('#download_status').html(getStatusHtml(actions.downloadStatus));
    $('#mailAll_request').html(REQUEST_HTML.replace('%TEXT%', actions.mailAllRequest));
    $('#mailAll_status').html(getStatusHtml(actions.mailAllStatus));
    $('#mailSpec_request').html(REQUEST_HTML.replace('%TEXT%', actions.mailSpecRequest));
    $('#mailSpec_status').html(getStatusHtml(actions.mailSpecStatus));
    $('#updateMailConfig_request').html(REQUEST_HTML.replace('%TEXT%', actions.updateMailConfigRequest));
    $('#updateMailConfig_status').html(getStatusHtml(actions.updateMailConfigStatus));
  });
};