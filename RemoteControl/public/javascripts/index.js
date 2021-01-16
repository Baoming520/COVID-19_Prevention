// Document ready event.
$(() => {
  // The home page has been loaded successfully.
});

$('#downloadBtn').on('click', (e) => {
  $.get('/actions', (actions) => {
    if(actions.downloadRequest === 'Yes'){
      alert('A download task is in progress now!');
      return;
    }

    window.href = '/download';
  });
});

$('#mailAllBtn').on('click', (e) => {
  var mailAddr = prompt("请输入合法的邮箱地址：\n");
  if (!mailAddr){
    alert('输入的内容为空，请重新输入！');
    return;
  }

  // Verify the email address using regex.
  // TODO:

  alert('口令将以邮件的形式发送至您的邮箱，请注意查收！');

  $.when($.get('/users/getGuestId')).then((guestId, textStatus, jqXHR) => {
    alert(guestId);
  });
  
  
  // var formData = new FormData();
  // formData.append('mailTo', mailAddr);
  // $.ajax({
  //   url: '/users/getToken',
  //   type: 'POST',
  //   async: false,
  //   processData: false,
  //   contentType: false,
  //   data: { mailTo: mailAddr },
  //   success: function (expToken) {
  //     var usrToken = prompt('请输入口令');
  //     if (usrToken === null) {
  //       return;
  //     }
  //     if (usrToken === expToken) {
  //       window.location.href = '/send_mail';
  //     }
  //     else {
  //       alert('口令错误，请重新操作');
  //     }
  //   },
  //   error: function (err) {
  //     alert(JSON.stringify(err));
  //   }
  // });
});

$('#mailSpecBtn').on('click', (e) => {
  var formData = new FormData();
  formData.append('token', '5oiR54ix5aSn54aK54yr');
  $.ajax({
    url: '/users/token',
    type: 'POST',
    async: false,
    processData: false,
    contentType: false,
    data: formData,
    success: function (pwd) {
      var token = prompt('请输入口令');
      if (token === null) {
        return;
      }
      if (token === pwd) {
        window.location.href = '/send_mail_spec';
      }
      else {
        alert('口令错误，请重新操作');
      }
    },
    error: function (err) {
      alert(JSON.stringify(err));
    }
  });
});

$('#updateMConfigBtn').on('click', (e) => {
  $('#msconf').trigger('click');
});

$('#msconf').on('change', (e) => {
  var files = e.target.files;
  if (files.length === 0) {
    return;
  }
  if (files.length > 1) {
    alert('仅可选择一个配置文件，请重新选择');
    return;
  }

  var msFile = files[0];
  if (!msFile.name.endsWith('.xlsx')) {
    alert('配置文件格式错误，请重新选择');
    return;
  }

  var formData = new FormData();
  formData.append('token', '5oiR54ix5aSn54aK54yr');
  $.ajax({
    url: '/users/token',
    type: 'POST',
    async: false,
    processData: false,
    contentType: false,
    data: formData,
    success: function (pwd) {
      var token = prompt('请输入口令');
      if (token === null) {
        return;
      }
      if (token === pwd) {
        $('#submit').trigger('click');
      }
      else {
        alert('口令错误，请重新操作');
      }
    },
    error: function (err) {
      alert(JSON.stringify(err));
    }
  });
});

var getToken = (redirectURL, callback) => {
  var formData = new FormData();
  formData.append('mailTo', mailAddr);
  $.ajax({
    url: '/users/getToken',
    type: 'POST',
    async: false,
    processData: false,
    contentType: false,
    data: formData,
    success: function (expToken) {
      var usrToken = prompt('请输入口令');
      if (usrToken === null) {
        return;
      }
      if (usrToken === expToken) {
        window.location.href = redirectURL;
      }
      else {
        alert('口令错误，请重新操作');
      }
    },
    error: function (err) {
      alert(JSON.stringify(err));
    }
  });
};