const OBSWebSocket = require('obs-websocket-js');

const sheetLoader = require('./sheet-loader');
const config = require('./config.json');

const getChildren = sources => {
	let items = sources;
	sources.forEach(source => {
		if (source.type === 'group') {
		items = items.concat(getChildren(source.groupChildren));
		}
	});
	return items;
}

const update = async (obs) => {
	
	const data = await sheetLoader.loadData();

	const { readFileSync } = require('fs');
	let json = readFileSync('./data.json', 'utf8');
	json = JSON.parse(json);
    	
	if ( data.toString() != json.toString()) {
				
		console.log("Sheets Updated");
  			
		const range = config.range;
		const startcell = range.split(":")[0].trim();

		const startcol = startcell.match("[a-zA-Z]+");
		//console.log("starting column is " + startcol);
		const startrow = startcell.match("[0-9]+");
		//console.log("starting row is " + startrow);

		const rowoffset = startrow[0];
		//console.log("row offset to array is " + rowoffset);
		const coloffset = columnToNumber(startcol[0]);
		//console.log("colum offset to array is " + coloffset);

		const sceneList = await obs.send('GetSceneList');
		await sceneList.scenes.forEach(async scene => {
			// unfold group children
			const allSources = getChildren(scene.sources);

			// console.log(scene);
			await allSources.forEach(async source => {
			  if (source.name.includes('|sheet')) {
				const reference = source.name.split('|sheet')[1].trim();

				let col = reference.match("[a-zA-Z]+");
				let colnumber = 3//columnToNumber(col[0]) - coloffset;
				
				let row = reference.match("[0-9]+");
				let rownumber = row[0] - rowoffset;

				let cellvalue = data[colnumber][rownumber];
				console.log("Value for cell in source is " + cellvalue)
				
				let sourcetype = data[2][rownumber]
				// If Source type is Text
				if (cellvalue != undefined) {
					
					if (sourcetype == "Text"){
						
						let color = null;

						if (cellvalue.startsWith('?color')) {
							const split = cellvalue.split(';');
							cellvalue = split[1];
							color = split[0].split('=')[1];
							color = color.replace('#', '');
							const color1 = color.substring(0, 2);
							const color2 = color.substring(2, 4);
							const color3 = color.substring(4, 6);
							color = parseInt('ff' + color3 + color2 + color1, 16);
						}

						if (cellvalue.startsWith('?hide')) {
							const split = cellvalue.split(';');
							cellvalue = split[1];
							await obs.send("SetSceneItemRender", {
								'scene-name': scene.name,
								source: source.name,
								render: false
							});
						} else if (cellvalue.startsWith('?show')) {
							const split = cellvalue.split(';');
							cellvalue = split[1];
							await obs.send("SetSceneItemRender", {
								'scene-name': scene.name,
								source: source.name,
								render: true
							});
						}
					  
					  
						// Update to OBS
						await obs.send("SetTextGDIPlusProperties", {
							source: source.name,
							text: cellvalue,
							color: color
						});
						console.log(`Updated: ${reference} to OBS: ${source.name}`);
						
					}
				
					// If Source type is Color
					if (sourcetype == "Color"){
						let color = null;
						color = cellvalue
						color = color.replace('#', '');
						const color1 = color.substring(0, 2);
						const color2 = color.substring(2, 4);
						const color3 = color.substring(4, 6);
						color = parseInt('ff' + color3 + color2 + color1, 16);
						
						await obs.send("SetSourceSettings", {
							sourceName: source.name,
							sourceType: source.type,
							sourceSettings: {
								color: color
							}
						});	
						
					}
					if (sourcetype == "Image"){
						
						let imagesettings = await obs.send("GetSourceSettings", {
							sourceName: source.name,
							sourceType: source.type,
						});	
						//console.log(imagesettings);
						let oldfile = imagesettings['sourceSettings']['file']
						//console.log(oldfile);
						if (cellvalue != oldfile){
							console.log('image updated');
							await obs.send("SetSourceSettings", {
								sourceName: source.name,
								sourceType: source.type,
								sourceSettings: {
									file: cellvalue
								}
							});	
						}			
					}
					if (sourcetype == "Browser"){
						
						await obs.send("SetBrowserSourceProperties", {
							source: source.name,
							url: cellvalue
						});	
						
					}
				}
			  }
			});
		});

		const fs = require('fs');
		const jsonContent = JSON.stringify(data);

		fs.writeFile("./data.json", jsonContent, 'utf8', function (err) {
			if (err) {
				return console.log(err);
			}
			console.log("The file was saved!");
		}); 
	}
}

const main = async () => {
	const obs = new OBSWebSocket();
	if (config.obsauth != "") {
		await obs.connect({ address: config.obsaddress, password: config.obsauth });
	}
	else {
		await obs.connect({ address: config.obsaddress });
	}
	console.log('Connected to OBS!');

	const updateWrapped = () => update(obs).catch(e => {
		console.log("EXECUTION ERROR IN MAIN LOOP:");
		console.log(e);
	});

	setInterval(updateWrapped, config.polling);
	updateWrapped();
}

main().catch(e => {
	console.log("EXECUTION ERROR:");
	console.log(e);
});

function columnToNumber(str) {
	var out = 0, len = str.length;
	for (pos = 0; pos < len; pos++) {
		out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
	}
	return out-1;
}