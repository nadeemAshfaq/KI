var app = angular.module('kiApp', ['ngMaterial', 'ngRoute']);



app.controller('kiController', function ($scope, $mdToast, $window, $mdDialog) {

    let baseapikey = window.localStorage.getItem("apiKey");
    baseapikey = JSON.parse(baseapikey)
    var apidata = ""
    if (baseapikey) {

        $scope.apiInput = true
        apidata = baseapikey
    } else {

        $scope.apiInput = false

    }

    Office.onReady(function () {

        var mailBody;

        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result) {


            if (result.status === Office.AsyncResultStatus.Succeeded) {
                mailBody = result.value;

               

            } else {
                console.error('Error retrieving mail body:', result.error.message);
            }
        });


        $scope.apiKey = '';
        $scope.companyInfo = '';

        $scope.companyInfo = 'your signature comes here'





        $scope.$apply();
        $scope.submitData = function () {

            var apiKeyValue = $scope.apiKey;

            if (apidata) {

            } else {

                window.localStorage.setItem("apiKey", JSON.stringify(apiKeyValue))
                apidata = $scope.apiKey;

            }

            var companyInfoValue = $scope.companyInfo;
            ProgressLinearActive()


            const API_URL = 'https://api.openai.com/v1/chat/completions';






            function bossApi() {

             
                const data = {
                    'model': 'gpt-3.5-turbo',
                    'messages': [
                        {
                            'role': 'system',
                            'content': 'you are very good assistant'
                        },
                        {
                            'role': 'user',
                            'content': 'This is the email I received {' + mailBody + '} suggest me the best answer for the email in the same language'
                        }
                    ],
                    'max_tokens': 1000,
                    'temperature': 0.5,
                    n: 1
                };

                return new Promise((resolve, reject) => {
                    $.ajax({
                        url: API_URL,
                        headers: {
                            'Authorization': `Bearer ${apidata}`,
                            'Content-Type': 'application/json'
                        },
                        method: 'POST',
                        dataType: 'json',
                        data: JSON.stringify(data),
                        success: function (response) {

                            const reply = response.choices[0].message.content;
                            $scope.apiInput = true
                            $scope.apiResponse = reply
                            ProgressLinearInActive()
                            resolve(reply);
                        },
                        error: function (jqXHR, textStatus, errorThrown) {
                            console.error('AJAX request failed:', textStatus, errorThrown);
                            
                            
                            ProgressLinearInActive()
                            loadToast("error please try again")

                            if (jqXHR.responseJSON && jqXHR.responseJSON.error) {
                                $window.localStorage.removeItem("apiKey");
                                $window.location.reload();
                            }
                            reject(new Error(errorThrown));
                        }
                    });
                });
            }
            bossApi()



        };

        function ProgressLinearActive() {
            $("#StartProgressLinear").show(function () {

                $("#ProgressBgDiv").show();
                $scope.ddeterminateValue = 15;
                $scope.showProgressLinear = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            });
        };
        function ProgressLinearInActive() {
            $("#StartProgressLinear").hide(function () {
                setTimeout(function () {
                    $scope.ddeterminateValue = 0;
                    $scope.showProgressLinear = true;
                    $("#ProgressBgDiv").hide();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }, 500);
            });
        };

        function loadToast(alertMessage) {
            var el = document.querySelectorAll('#zoom');
            $mdToast.show(
                $mdToast.simple()
                    .textContent(alertMessage)
                    .position('bottom')
                    .hideDelay(4000))
                .then(function () {
                    $log.log('Toast dismissed.');
                }).catch(function () {
                    $log.log('Toast failed or was forced to close early by another toast.');
                });
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        };

    });
});