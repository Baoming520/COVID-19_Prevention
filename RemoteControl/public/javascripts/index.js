// Document ready event.
$(document).ready(function () {
});

$('#sendMailBtn').click(function (e) {
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
        window.location.href = '/send_mail';
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

$('#sendMailSpecBtn').click(function (e) {
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

$('#updateMConfigBtn').click(function (e) {
  $('#msconf').trigger('click');
});

$('#msconf').change(function (e) {
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
})