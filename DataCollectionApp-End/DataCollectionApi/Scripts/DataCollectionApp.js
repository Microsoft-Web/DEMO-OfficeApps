Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#save').click(function () {
            var url = 'http://localhost:31563/api/datacollection';
            Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, { id: 'myBinding' }, function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    write('Action failed. Error: ' + asyncResult.error.message);
                } else {
                    asyncResult.value.getDataAsync({ coercionType: 'matrix' }, function (e) {

                        var j = $.map(e.value, function (val, y) {
                            return { first: val[0], last: val[1] };
                        });

                        $.each(j, function (i, item) {
                            $.post(url, item, function (data) {
                                $('#tbl').append('<tr><td>' + data + '</td></tr>');
                            });
                        });
                    });
                }
            });
        });
    });
};