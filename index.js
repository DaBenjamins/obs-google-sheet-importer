const OBSWebSocket = require('obs-websocket-js');

const sheetLoader = require('./sheet-loader');
const config = require('./config.json');

let json = [];

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
	
	//read data.json
	const { readFileSync } = await require('fs');
	json = readFileSync('./data.json', 'utf8');
	if (json == undefined){
		json = [];
	}
	json = JSON.parse(json);
    
	//check if sheets is same as json
	if ( data.toString() != json.toString()) {
		
		console.log("Sheets Updated");		
		
		// Write data.json to check if sheets been changed
		const fs = await require('fs');
		const jsonContent = await JSON.stringify(data);
		await fs.writeFile("./data.json", jsonContent, 'utf8', function (err) {
			if (err) {
				return console.log(err);
			}
			console.log("The file was saved!");
		});
			
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
		
		await Promise.all(
		await sceneList.scenes.map(async scene => {
			// unfold group children
			const allSources = getChildren(scene.sources);

			// console.log(scene);
			await Promise.all(
			await allSources.map(async source => {
				if (source.name.includes('|sheet')) {
					const reference = source.name.split('|sheet')[1].trim();
					
					let row = reference.match("[0-9]+");
					let rownumber = row[0] - rowoffset;

					let cellvalue = data[1][rownumber];
					let sourcetype = data[0][rownumber]
					console.log("Value for cell in source is " + cellvalue + " and sourcetype " + sourcetype)
					
					// If Cell is empty skip
					if (cellvalue != undefined) {
						// If Source type is Text
						if (sourcetype == "Text"){
							let color = null;
							//check if ?color tag is present
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
							//check if ?hide/?show tag is present
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
							//get settings of source from OBS
							let textsettings = await obs.send("GetTextGDIPlusProperties", {
								source: source.name
							});
							let oldfile = await textsettings['text']
							let oldcolor = await textsettings['color']
							//check if current OBS settings is different
							if (cellvalue != oldfile){
								if (color == null){
									color = oldcolor
								}
								// Update to OBS
								await obs.send("SetTextGDIPlusProperties", {
									source: source.name,
									text: cellvalue,
									color: color
								});
								console.log(`Updated: ${reference} to OBS: ${source.name}`);
							} else {
								console.log('text is the same');
							}
							
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
							//get settings of source from OBS
							let colorsettings = await obs.send("GetSourceSettings", {
								sourceName: source.name,
								sourceType: source.type,
							});
							let oldfile = await colorsettings['sourceSettings']['color']
							//check if current OBS settings is different
							if (color != oldfile){
								console.log(`Updated: ${reference} to OBS: ${source.name}`);
								await obs.send("SetSourceSettings", {
									sourceName: source.name,
									sourceType: source.type,
									sourceSettings: {
										color: color
									}
								});	
							} else {
								console.log('Color is the same');
							}
						}
						// If Source type is Image
						if (sourcetype == "Image"){
							//get settings of source from OBS
							let imagesettings = await obs.send("GetSourceSettings", {
								sourceName: source.name,
								sourceType: source.type,
							});	
							let oldfile = await imagesettings['sourceSettings']['file']
							//check if current OBS settings is different
							if (cellvalue != oldfile){
								console.log(`Updated: ${reference} to OBS: ${source.name}`);
								await obs.send("SetSourceSettings", {
									sourceName: source.name,
									sourceType: source.type,
									sourceSettings: {
										file: cellvalue
									}
								});	
							} else {
								console.log('Image is the same');
							}
						}
						// If Source type is Browser
						if (sourcetype == "Browser"){
							//get settings of source from OBS
							let browsersettings = await obs.send("GetBrowserSourceProperties", {
								source: source.name
							});
							let oldfile = await browsersettings['url']
							//check if current OBS settings is different
							if (cellvalue != oldfile){
								console.log(`Updated: ${reference} to OBS: ${source.name}`);
								await obs.send("SetBrowserSourceProperties", {
									source: source.name,
									url: cellvalue
								});
							} else {
								console.log('Browser is the same');
							}
						}
						// If Source type is HS
						if (sourcetype == "HS"){
							if (cellvalue.startsWith('hide')) {
								await obs.send("SetSceneItemRender", {
									'scene-name': scene.name,
									source: source.name,
									render: false
								});
							} else if (cellvalue.startsWith('show')) {
								await obs.send("SetSceneItemRender", {
									'scene-name': scene.name,
									source: source.name,
									render: true
								});
							}
							console.log(`Updated: ${reference} to OBS: ${source.name}`);
						}
					}
				}
			}));
		})); 
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

update().catch(e =>  {
    console.log('------------------');
    console.log(e.message);
    console.log('--------')
})

function columnToNumber(str) {
	var out = 0, len = str.length;
	for (pos = 0; pos < len; pos++) {
		out += (str.charCodeAt(pos) - 64) * Math.pow(26, len - pos - 1);
	}
	return out-1;
}