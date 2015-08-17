var http = require("http"),
	officegen = require("officegen"),
	syncrequest = require("sync-request"),
	sizeOf = require("image-size"),
	path = require("path"),
	url = require("url");

var apikey = process.env.APIKEY;
var port = process.env.PORT || 5000;

http.createServer(function(request, response) {
	var parsedUrl = url.parse(request.url, true); // true to get query as object
  	var queryString = parsedUrl.query;
  	var objectIDList = queryString.objectid.split(",");

	response.writeHead(200, {
		"Content-Type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
		"Content-disposition": "attachment;filename=slides.pptx"
	});

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
		//Get the object record
		var output = syncrequest('GET', 'http://api.harvardartmuseums.org/object/' + objectid + "?apikey=" + apikey);
		var object = JSON.parse(output.getBody());		

		if (!object.error) {
			//Get the image of the object
			var imageParameters = "width=840&height=500";
			var image, imageDims;
			if (object.imagepermissionlevel > 0) {
				imageParameters = "width=256&height=256";
			}
			if (object.images[0]) {
				image = syncrequest('GET', "http://ids.lib.harvard.edu/ids/view/" + object.images[0].idsid + "?" + imageParameters);
				imageDims = sizeOf(image.getBody());
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

			slide = pptx.makeNewSlide();
			slide.back = 'ffffff';
			slide.color = '000000';
			slide.name = object.objectnumber;
			if (image) {
				slide.addImage(image.getBody(), {y: 20, x: 'c', cx: imageDims.width, cy: imageDims.height});
				textVerticalOffset += imageDims.height
			}
			slide.addText(description, {y: textVerticalOffset, cx: '100%', font_size: 10, font_face: 'Arial', bold: false, color: '000000', align: 'left'});
		}
	});

	pptx.generate(response);
}).listen(port);

console.log("Listening on 3000...");