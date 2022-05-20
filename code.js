(function (window, undefined) {
    window.Asc.plugin.init = function () {
        this.callCommand(function() {
            const sheet = Api.GetActiveSheet();
            const range = sheet.GetUsedRange();
            const color = Api.CreateColorFromRGB(255, 224, 204);
            const newRowIndex = range.GetRows().GetCount() + 1;
        
            const setValue = (selector, value) => sheet
                .GetRange(selector)
                .SetValue(value);
                            
            setValue('A' + newRowIndex, 'plugin');
            setValue('B' + newRowIndex, 299);
            setValue('C' + newRowIndex, 500);        
        }, true);
    };
})(window, undefined);