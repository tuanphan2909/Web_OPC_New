


$(document).ready(function() {
    $('#addRowBtn').click(function() {
        var newRow = '<tr>' +
            '<td><input type="text" name="Ma_vt" /></td>' +
            '<td><input type="text" name="Ten_vt" /></td>' +
            '<td><input type="text" name="Han_Muc" /></td>' +
            '<td><button class="deleteRowBtn">Xóa</button></td>' +
            '</tr>';
        $('#example tbody').append(newRow);
    });

    $(document).on('click', '.deleteRowBtn', function() {
        $(this).closest('tr').remove();
    });
});