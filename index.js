
process.env['PATH'] = process.env['PATH'] + ':' + process.env['LAMBDA_TASK_ROOT'];
console.log('Loading function');

var AWS = require("aws-sdk"),
	officegen = require("officegen"),
	syncrequest = require("sync-request"),
	sizeOf = require("image-size"),
	path = require("path"),
	fs = require("fs"),
	url = require("url");

var apikey = "7d11a0a0-4522-11e5-806d-cd538555eccd";


exports.handler = (event, context, callback) => {

	var queryString = event.objectid;
	var collectionid = event.collectionid;
	var useremail = event.useremail;
    
    if (!queryString === "") {
  		callback("No objects supplied");

  	} else {
	  	var objectIDList = queryString.split(",");

		var pptx = officegen("pptx");
		pptx.on('finalize', function(written) {

		});
		pptx.on('error', function(error) {

		});

		var slide;

		//Title slide
		slide = pptx.makeNewSlide();	
		slide.back = 'ffffff';
		slide.color = 'ffffff';
		slide.addImage(path.resolve(__dirname, 'static/logo.jpg'), {y: 0, x: 0, cx: 750, cy: 300});

		//Work through the object list making a slide for each
		objectIDList.forEach(function(objectid) {
			if (!isNaN(objectid) && objectid !== "") {
				//Get the object record
				var output = syncrequest('GET', 'http://api.harvardartmuseums.org/object/' + objectid + "?apikey=" + apikey);
				var object = JSON.parse(output.getBody());		

				if (!object.error) {
					//Get the image of the object
					var imageParameters = "width=840&height=500";
					var image, imageDims, imageCaption;
					if (object.imagepermissionlevel > 0) {
						imageParameters = "width=256&height=256";
						imageCaption = "(Large image restricted)";
					}
					if (object.images) {
						if (object.images[0]) {
							image = syncrequest('GET', object.images[0].baseimageurl + "?" + imageParameters);
							//Add some error handling
							//Check to make sure image.headers["content-type"] === "image/jpeg
							//If not, use a place holder image indicating an error has occurred
							imageDims = sizeOf(image.getBody());
						}
					}

					//Build the object data in to a description text block
					var description = "";
					if (object.titles) {
						object.titles.forEach(function(title) {
							description += title.titletype + ": " + title.title + "\n";
						});
					}
					if (object.people) {
						object.people.forEach(function(person) {
							description += person.role + ": " + person.displayname + "\n";
						});
					}
					if (object.classification) {
						description += "Classification: " + object.classification + "\n";
					}
					if (object.worktypes) {
						var workTypes = "";
						for (var i = 0; i < object.worktypes.length; i++) {
							workTypes += object.worktypes[i].worktype;
							if (i < object.worktypes.length-1) workTypes += ", ";
						};
						description += "Work Type: " + workTypes + "\n";
					}
					if (object.culture) {
						description += "Culture: " + object.culture + "\n";
					}
					if (object.places) {
						object.places.forEach(function(place) {
							description += place.type + ": " + place.displayname + "\n";
						});
					}
					if (object.dated) {
						description += "Date: " + object.dated + "\n";
					}
					if (object.century) {
						description += "Century: " + object.century + "\n";
					}
					if (object.period) {
						description += "Period: " + object.period + "\n";
					}
					if (object.medium) {
						description += "Medium: " + object.medium + "\n";
					}
					if (object.technique) {
						description += "Technique: " + object.technique + "\n";
					}
					if (object.dimensions) {
						description += "Dimensions: " + object.dimensions.replace("\r\n", "; ") + "\n";
					}
					description += "Credit line: " + object.creditline + "\n";
					description += "Object number: " + object.objectnumber + "\n";
					if (object.copyright) {
						description += "Copyright: " + object.copyright + "\n";
					}
					description += "Division: " + object.division + "\n";
					description += "URL: " + object.url + "\n";

					//Put the image and description in the slide
					var textVerticalOffset = 40;
					var marginTop = 20;

					slide = pptx.makeNewSlide();
					slide.back = 'ffffff';
					slide.color = '000000';
					slide.name = object.objectnumber;
					if (image) {
						slide.addImage(image.getBody(), {y: marginTop, x: 'c', cx: imageDims.width, cy: imageDims.height});
						slide.addText(imageCaption, {y: marginTop + imageDims.height, x: 'c', cx: '100%', font_size: 8, font_face: 'Arial', color: '000000', align: 'center'});
						textVerticalOffset += imageDims.height
					}
					slide.addText(description, {y: textVerticalOffset, cx: '100%', font_size: 10, font_face: 'Arial', bold: false, color: '000000', align: 'left'});
				}
			}
		});
	var out = fs.createWriteStream ( '/tmp/out.pptx' );
	pptx.generate ( out );
	out.on ( 'close', function () {
  		console.log ( 'Finished creating the PPTX file!' );
		var s3 = new AWS.S3();
    	var pptx_local = fs.readFileSync('/tmp/out.pptx');
    	var filename = collectionid + ".pptx";
    	var param = {Bucket: 'slides.harvardartmuseums.org', Key: filename, Body:pptx_local, ACL:'public-read'};
    	console.log("s3");
    	var JSONpayload = new Array();
    	s3.upload(param, function(err, data) {
        	if (err) console.log(err, err.stack); // an error occurred
        	else console.log(data);           // successful response
    	});
    console.log('done');

    callback(null, {"location": "https://s3.amazonaws.com/slides.harvardartmuseums.org/" . filename});
		});
	}
 
 	
};