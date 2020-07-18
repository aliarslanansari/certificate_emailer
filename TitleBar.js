const customTitlebar = require('custom-electron-titlebar');
    let MyTitleBar = new customTitlebar.Titlebar({
        backgroundColor: customTitlebar.Color.fromHex('#444'),
        shadow: true,
        icon: './assets/icons/png/icon.png'
    });
MyTitleBar.updateTitle('');