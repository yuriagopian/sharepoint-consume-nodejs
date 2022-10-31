const { SPPull } = require("sppull");

var context = {
  siteUrl: "https://{seller}.sharepoint.com/",
  username: "",
  password: "",
};

var options = {
  spRootFolder: "",
  dlRootFolder: "./Downloads/Contracts",
};


SPPull.download(context, options)
  .then((downloadResults) => {
    console.log("Files are downloaded");
    console.log("For more, please check the results", JSON.stringify(downloadResults));
  })
  .catch((err) => {
    console.log("Core error has happened", err);
    // error: https://docs.microsoft.com/en-us/answers/questions/646771/34access-has-been-blocked-by-conditional-access-po.html
  });

  function capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
  }