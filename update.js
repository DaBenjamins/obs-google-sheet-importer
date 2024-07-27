const AutoGitUpdate = require('auto-git-update');

const configUP = {
    repository: 'https://github.com/DaBenjamins/obs-google-sheet-importer',
    fromReleases: true,
    tempLocation: './tmp/',
}

const updater = new AutoGitUpdate(config);

const versionComp = updater.compareVersions();
console.log(versionComp);
	
updater.forceUpdate();
