var app = angular.module('MyApp', ['ngFileUpload'])
app.controller('MyController', function ($scope, $window) {
    $scope.SelectFile = function (file) {
        $scope.SelectedFile = file;
        $scope.Upload();
    };
    $scope.Upload = function () {
        var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
        if (regex.test($scope.SelectedFile.name.toLowerCase())) {
            if (typeof (FileReader) != "undefined") {
                var reader = new FileReader();
                //For Browsers other than IE.
                if (reader.readAsBinaryString) {
                    reader.onload = function (e) {
                        $scope.ProcessExcel(e.target.result);
                    };
                    reader.readAsBinaryString($scope.SelectedFile);
                } else {
                    //For IE Browser.
                    reader.onload = function (e) {
                        var data = "";
                        var bytes = new Uint8Array(e.target.result);
                        for (var i = 0; i < bytes.byteLength; i++) {
                            data += String.fromCharCode(bytes[i]);
                        }
                        $scope.ProcessExcel(data);
                    };
                    reader.readAsArrayBuffer($scope.SelectedFile);
                }
            } else {
                $window.alert("This browser does not support HTML5.");
            }
        } else {
            $window.alert("Please upload a valid Excel file.");
        }
    };

    $scope.ProcessExcel = function (data) {
        //Read the Excel File data.
        var workbook = XLSX.read(data, {
            type: 'binary'
        });

        //Fetch the name of First Sheet.
        var firstSheet = workbook.SheetNames[0];

        //Read all rows from First Sheet into an JSON array.
        $scope.excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
        //Display the data from Excel file in Table.
        $scope.$apply(function () {
            $scope.goodData = $scope.excelRows.filter(m => m.isRefQK == 1);
            $scope.badData = $scope.excelRows.filter(m => m.isRefQK == 0);
            $scope.allData = $scope.excelRows.filter(m => m.isTitleRefRelevant == 1 && m.IsTitleRelevant == 1);
            $scope.dataAnalysis();
            $scope.IsVisible = true;
        });
    };
    $scope.dataAnalysis = function () {
        $scope.rotten = 0;
        $scope.normal = 0;
        $scope.indexTest = 0;
        $scope.referenceType = 0;
        $scope.wordsToCheck = {
            firstGroop: {
                one: ['סימן', 'סימנים', 'פרט', 'סעיפים', 'סעיף', 'תקנה', 'תקנות'],
                two: ['פרק', 'חלק']
            },
            secondGroop: {
                one: ['פקודה', 'פקודת'],
                two: ['חוק'],
                three: ['תוספות', 'תוספת'],
                four: ['כללי']
            }
        };
        for (var i = 0; i < $scope.excelRows.length; i++) {
            $scope.reference = $scope.excelRows[i];
            if ($scope.reference.isTitleRefRelevant != 1 || $scope.reference.IsTitleRelevant != 1 || $scope.reference.ref.includes('http&#58;//')) {
                $scope.excelRows[i].result = 0;
                $scope.indexTest++;
            }
            else {
                $scope.firstGroupCompatibilityCheck($scope.wordsToCheck.firstGroop.one, $scope.reference, i);
                $scope.firstGroupCompatibilityCheck($scope.wordsToCheck.firstGroop.two, $scope.reference, i);
                $scope.secondGroupCompatibilityCheck($scope.wordsToCheck.secondGroop.one, $scope.reference, i);
                $scope.secondGroupCompatibilityCheck($scope.wordsToCheck.secondGroop.two, $scope.reference, i);
                $scope.secondGroupCompatibilityCheck($scope.wordsToCheck.secondGroop.three, $scope.reference, i);
                $scope.secondGroupCompatibilityCheck($scope.wordsToCheck.secondGroop.four, $scope.reference, i);
            }
            if (!$scope.excelRows[i].result) {
                $scope.excelRows[i].result = 0;
            }

        }
        $scope.rotten = $scope.excelRows.filter(e => e.result == 0 && e.isRefQK >= 0).length;
        $scope.normal = $scope.excelRows.filter(e => e.result == 1 && e.isRefQK >= 0).length;
        //homan
        $scope.goodDataInRefernce = $scope.excelRows.filter(m => m.isRefQK == 1 && m.isSameLawRef == 1).length;
        $scope.goodDataOutRefernce = $scope.excelRows.filter(m => m.isRefQK == 1 && m.isSameLawRef == 0).length;
        $scope.badDataInRefernce = $scope.excelRows.filter(m => m.isRefQK == 0 && m.isSameLawRef == 1).length;
        $scope.badDataOutRefernce = $scope.excelRows.filter(m => m.isRefQK == 0 && m.isSameLawRef == 0).length;
        //mashin
        $scope.normalInRefernce = $scope.excelRows.filter(m => m.isRefQK >= 0 && m.isSameLawRef == 1 && m.result == 1).length;
        $scope.normalOutRefernce = $scope.excelRows.filter(m => m.isRefQK >= 0 && m.isSameLawRef == 0 && m.result == 1).length;
        $scope.rottenInRefernce = $scope.excelRows.filter(m => m.isRefQK >= 0 && m.isSameLawRef == 1 && m.result == 0).length;
        $scope.rottenOutRefernce = $scope.excelRows.filter(m => m.isRefQK >= 0 && m.isSameLawRef == 0 && m.result == 0).length;

    };
    $scope.firstGroupCompatibilityCheck = function (groop, reference, index) {
        if (!reference.result) {
            var temp1, temp2;
            for (var i = 0; i < groop.length; i++) {
                if (reference.ref.includes(groop[i])) {
                    if (reference.ref.includes(groop[i] + '_'))
                        temp2 = reference.ref.split(groop[i] + '_');
                    else temp2 = reference.ref.split(groop[i] + ' ');
                    temp2 = temp2[temp2.length - 1];
                    if (temp2.includes('('))
                        temp2 = temp2.split('(')[0];
                    else if (temp2.includes('.'))
                        temp2 = temp2.split('.')[0];
                    else if (temp2.includes(' '))
                        temp2 = temp2.split(' ')[0];
                    if (reference.title.includes(temp2))
                        $scope.excelRows[index].result = 1;
                }
            }
        }

    };
    $scope.secondGroupCompatibilityCheck = function (groop, reference, index) {
        if (!reference.result) {
            var temp1, temp2;
            for (var i = 0; i < groop.length; i++) {
                if (reference.ref.includes(groop[i])) {
                    for (var j = 0; j < groop.length; j++) {
                        if (reference.title.includes(groop[j])) {
                            temp1 = reference.title.split(groop[j] + ' ');
                            temp1 = temp1[temp1.length - 1].split(',')[0];
                            if (temp1.includes('[צ"ל:')) {
                                temp1 = temp1.split('[צ"ל:');
                                if (temp1.length > 1)
                                    temp1 = temp1[0] + temp1[1];
                                temp1 = temp1.split(']');
                                if (temp1.length > 1)
                                    temp1 = temp1[0] + temp1[1];
                            }
                            temp2 = reference.ref.split(groop[i] + ' ');
                            temp2 = temp2[temp2.length - 1];
                            if (temp2.includes('('))
                                temp2.split('(')[0];
                            else if (temp2.includes('.'))
                                temp2.split('.')[0];
                            if (temp2 == temp1) {
                                j = groop.length;
                                i = groop.length;
                                $scope.excelRows[index].result = 1;
                            }

                        }
                    }
                }
            }
        }

    };
    $scope.createNewFile = function () {
        alasql('SELECT * INTO XLSX("output.xlsx",{headers:true}) FROM ?', [$scope.excelRows]);
    };
});