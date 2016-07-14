var gulp = require('gulp'),
  sharepoint = require('./gulp-sharepoint'),
  constants = require('./config/O365Constants.js'),
  args=require('yargs').argv;

gulp.task('deploy-sharepoint', function () {
  console.log(constants.resource);
  sharepoint.getToken(constants.resource, constants.clientId, constants.clientSecret, function (token) {
    sharepoint.getCustomAction(constants.resource,null,token, function (actions) {
      customactions=actions.d.results;
      var index=0;
      customactions.forEach(function(element) {
        if(element.Name===args.Name){
          sharepoint.deleteCustomAction(constants.resource,element.Id,token,function(res){
            index++;            
            console.log('action '+element.Id +' has been deleted');
            if(index==customactions.length){
              createCustomAction(token);
            }
          });
        }
      }, this);
      if(customactions.length==0){
        createCustomAction(token);        
      }
    });
  });
});

function createCustomAction(token){
  var customAction={};
  customAction.Name=args.Name;
  customAction.Location= args.Location;
  customAction.RegistrationId=args.RegistrationId+'';
  customAction.RegistrationType=args.RegistrationType+'';
  customAction.Sequence=args.Sequence+'';
  customAction.Title= args.Title;
  customAction.Url=args.Url;
  customAction.Group='NodeJS Group';
  customAction.Description='NodeJS Custom user action example.';  
  sharepoint.createCustomAction(constants.resource,customAction,token,function(res){
    console.log('Custom action has been created succesfully');
  });
}

