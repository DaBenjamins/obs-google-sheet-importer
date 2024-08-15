const {default: OBSWebSocket} = require('obs-websocket-js');

const sheetLoader = require('./sheet-loader');
const config = require('./config.json');

var json = JSON.parse("[]");

const update = async (obs) => {
	const data = await sheetLoader.loadData();
	
	// check if sheets is same as json
	if ( data.toString() != json.toString()) {
		json = data
		
		console.log("Sheets Updated");		
		
		const range = config.range;
		const startcell = range.split(":")[0].trim();
		const startrow = startcell.match("[0-9]+");
		const rowoffset = startrow[0];
		const sceneList = await obs.call('GetSceneList');
		const groupList = await obs.call('GetGroupList');
		var sceneNameList = []
		var allSources
		
		// Gets all Scene and Group names
		await sceneList.scenes.forEach(async scene => {
			sceneNameList.push([scene.sceneName, 0]);
		});
		await groupList.groups.forEach(async group => {
			sceneNameList.push([group, 1])
		});
		
		await sceneNameList.forEach(async scene => {
			// unfold group children
			
			// If Group
			if(scene[1] == 0){
				allSources = await obs.call('GetSceneItemList', {sceneName: scene[0]});
			} else {
				allSources = await obs.call('GetGroupSceneItemList', {sceneName: scene[0]});
			}

			await allSources.sceneItems.forEach(async source => {
				if (source.sourceName.includes('|sheet')) {
					const reference = source.sourceName.split('|sheet')[1].trim();
					let row = reference.match("[0-9]+");
					let rownumber = row[0] - rowoffset;
					let cellvalue = data[1][rownumber];
					let sourcetype = data[0][rownumber];
		
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
							await obs.call("SetSceneItemEnabled", {
								sceneName: scene[0],
								sceneItemId: source.sceneItemId,
								sceneItemEnabled: false
							});
						} else if (cellvalue.startsWith('?show')) {
							const split = cellvalue.split(';');
							cellvalue = split[1];
							await obs.call("SetSceneItemEnabled", {
								sceneName: scene[0],
								sceneItemId: source.sceneItemId,
								sceneItemEnabled: true
							});
						}
						//get settings of source from OBS
						let textsettings = await obs.call("GetInputSettings", {
							inputName: source.sourceName
						});
						let oldfile = await textsettings['inputSettings']['text']
						let oldcolor = await textsettings['inputSettings']['color']
						//check if current OBS settings is different
						if (cellvalue != oldfile){
							if (color == null){
								color = oldcolor
							}
							// Update to OBS
							await obs.call("SetInputSettings", {
								inputName: source.sourceName,
								inputSettings: {
									text: cellvalue,
									color: color
								}
							});
							console.log(`Updated: ${reference} from ${oldfile} to ${cellvalue} on source: ${source.sourceName}`);
						} else {
							//console.log('text is the same');
						}
					}
					// If Source type is Color
					if (sourcetype == "Color"){
						if (cellvalue != undefined) {
							let color = null;
							color = cellvalue
							color = color.replace('#', '');
							const color1 = color.substring(0, 2);
							const color2 = color.substring(2, 4);
							const color3 = color.substring(4, 6);
							color = parseInt('ff' + color3 + color2 + color1, 16);
							//get settings of source from OBS
							let colorsettings = await obs.call("GetInputSettings", {
								inputName: source.sourceName,
							});
							let oldfile = await colorsettings['inputSettings']['color']
							//check if current OBS settings is different
							if (color != oldfile){
								console.log(`Updated: ${reference} from ${oldfile} to ${color} on source: ${source.sourceName}`);
								await obs.call("SetInputSettings", {
									inputName: source.sourceName,
									inputSettings: {
										color: color
									}
								});	
							} else {
								//console.log('Color is the same');
							}
						}
					}
					// If Source type is Image
					if (sourcetype == "Image"){
						let permvalue = ""
						//get settings of source from OBS
						let imagesettings = await obs.call("GetInputSettings", {
							inputName: source.sourceName,
						});	
						let oldfile = await imagesettings['inputSettings']['file']
						let hidetime = 1;
						if (cellvalue != undefined) {
							if (cellvalue.startsWith('?')) {
								const split = cellvalue.split(';');
								cellvalue = split[1];
								permvalue = split[0];
							}
						}
						//Hide
						if (permvalue.startsWith('?hide')) {
							obs.call("SetSceneItemEnabled", {
								sceneName: scene[0],
								sceneItemId: source.sceneItemId,
								sceneItemEnabled: false
							});
							hidetime = 1500
						}
						//check if current OBS settings is different
						setTimeout(function(){
							if (cellvalue != oldfile){
								console.log(`Updated: ${reference} from ${oldfile} to ${cellvalue} on source: ${source.sourceName}`);
								obs.call("SetInputSettings", {
									inputName: source.sourceName,
									inputSettings: {
										file: cellvalue
									}
								});	
							} else {
								//console.log('Image is the same');
							}
						}, hidetime);
						//Show
						setTimeout(function(){
							if (permvalue.startsWith('?show')) {
								obs.call("SetSceneItemEnabled", {
									sceneName: scene[0],
									sceneItemId: source.sceneItemId,
									sceneItemEnabled: true
								});
							}
						}, 750);
					}
					// If Source type is Browser
					if (sourcetype == "Browser"){
						//get settings of source from OBS
						let browsersettings = await obs.call("GetInputSettings", {
							inputName: source.sourceName
						});
						let oldfile = await browsersettings['inputSettings']['url']
						//check if current OBS settings is different
						if (cellvalue != oldfile){
							console.log(`Updated: ${reference} from ${oldfile} to ${cellvalue} on source: ${source.sourceName}`);
							await obs.call("SetInputSettings", {
								inputName: source.sourceName,
								inputSettings: {
									url: cellvalue
								}
							});
						} else {
							//console.log('Browser is the same');
						}
					}
					// If Source type is HS
					if (sourcetype == "HS"){
						if (cellvalue.startsWith('hide')) {
							await obs.call("SetSceneItemEnabled", {
								sceneName: scene[0],
								sceneItemId: source.sceneItemId,
								sceneItemEnabled: false
							});
							console.log(`Updated: ${reference} set to hidden on source: ${source.sourceName}`);
						} else if (cellvalue.startsWith('show')) {
							await obs.call("SetSceneItemEnabled", {
								sceneName: scene[0],
								sceneItemId: source.sceneItemId,
								sceneItemEnabled: true
							});
							console.log(`Updated: ${reference} set to visible on source: ${source.sourceName}`);
						}
					}
				}
			});  
		});

	}
}

const main = async () => {
	const obs = new OBSWebSocket();
	let coned = 0

	await obs.connect(config.obsaddress, config.obsauth).catch(e => {
		console.log("FAILED TO CONNECT, Reconnecting in 5 seconds");
		coned = 1
		setTimeout(function(){
		  main();
		}, 5000);
	});

	if (coned == 0){
		
		console.log('Connected to OBS!');

		const updateWrapped = () => update(obs).catch(e => {
			if(e.message == "Not connected"){
				obs.connect(config.obsaddress, config.obsauth).catch(e => {
					console.log("FAILED TO RECONNECT");
				});
			}else{
			console.log("EXECUTION ERROR IN MAIN LOOP:");
			console.log(e);
			}
		});
		setInterval(updateWrapped, config.polling);
	}

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