function toggleNumberInput() {
    var crawlOption = document.getElementById('crawl_option').value;
    var numberInputDiv = document.getElementById('number_input_div');
    if (crawlOption === 'direct') {
        numberInputDiv.style.display = 'block';
    } else {
        numberInputDiv.style.display = 'none';
    }
}

function showLoading() {
    document.getElementById("loading-overlay").style.display = "flex";
}

function hideLoading() {
    document.getElementById("loading-overlay").style.display = "none";
}

function showCompleted() {
    document.getElementById("completed-overlay").style.display = "flex";
}

function hideCompleted() {
    document.getElementById("completed-overlay").style.display = "none";
}

function submitForm(event, actionValue) {
    event.preventDefault();

    var username = document.getElementById('username').value.trim();
    var password = document.getElementById('password').value.trim();

    if (!username || !password) {
        alert('아이디와 비밀번호를 모두 입력해주세요.');
        return;
    }

    showLoading();

    var form = document.getElementById('data-form');
    var formData = new FormData(form);
    formData.append('action', actionValue);

    fetch('/', {
        method: 'POST',
        body: formData
    })
    .then(response => response.text())
    .then(data => {
        hideLoading();
        showCompleted();
        setTimeout(() => {
            hideCompleted();
            document.open();
            document.write(data);
            document.close();
        }, 3000);
    })
    .catch(error => {
        hideLoading();
        alert('오류가 발생했습니다: ' + error);
    });
}

window.onload = function() {
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    document.getElementById('max_posts').value = '';
};
