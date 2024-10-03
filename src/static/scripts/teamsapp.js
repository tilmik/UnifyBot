(function () {
  "use strict";

  // Call the initialize API first
  microsoftTeams.app.initialize().then(function () {
    microsoftTeams.app.getContext().then(function (context) {
      if (context?.app?.host?.name) {
        updateHubState(context.app.host.name);
        updateTheme(context.app.theme);
      }
    });
  });

  function updateHubState(hubName) {
    if (hubName) {
      document.getElementById("hubState").innerHTML = "Your app is running in " + hubName;
    }
  }

  function updateTheme(themeName) {
    if (themeName) {
      //document.getElementById("themeName").innerHTML = "Selected theme is " + themeName;
      if (themeName == "dark") {
        document.body.classList.toggle("dark-mode");
      } else if (themeName == "contrast") {
        document.body.classList.toggle("contrast-mode");
      }
    }
  }
})();
