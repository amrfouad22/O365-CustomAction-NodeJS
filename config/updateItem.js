var O365Connect = require('./O365Connect.js');

var express=require('express');
var router=express.Router();
router.get('/',function(req,res,next){
    O365Connect.updateItemTitle(req.query.SPListId.replace('{','').replace('}',''),req.query.SPListItemId,function(){
        res.end('<h1>Title has been updated</h1>');
    });
});
module.exports=router;