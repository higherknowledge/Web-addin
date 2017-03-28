/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  };


  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
      var item = Office.context.mailbox.item;
      var address = Office.context.mailbox.userProfile.emailAddress;
     
      $("#intern").click(handleIntern);
      $("#recruit").click(handleRecruit);

    function handleIntern()
      {
        
        if(localStorage["intern"] == null)
        {
            fetchTemp(1);
        }
        else
        {
            send(localStorage["intern"]);
        }
        
    }

    

    function handleRecruit()
    {
        if(localStorage["recruit"] == null)
        {
            fetchTemp(2);
        }
        else
        {
            send(localStorage["recruit"]);
        }
        
    }

  }

  function fetchTemp(flag) {
      var oReq = new XMLHttpRequest();
      
      function reqListener() 
      {
          showNotification("response", oReq.status);
          if(oReq.readyState == oReq.DONE && oReq.status == 200)
          {
            if (flag == 1)
                localStorage["intern"] = this.responseText;
            else
                localStorage["recruit"] = this.responseText;
            showNotification("Loading", "fetched");
            send(this.responseText);
          }
      }

      oReq.onreadystatechange = reqListener;
      oReq.open("GET", "https://raw.githubusercontent.com/higherknowledge/outlook-integration/master/templates/" + Office.context.mailbox.userProfile.emailAddress + (flag == 1 ? "" : "R"));
      oReq.send();
      showNotification("Loading", "fetching template...");
  }

  function send(template) {
      var response = JSON.parse(template);
      var body = getBody(response["Body"]);
      Office.context.mailbox.item.displayReplyForm(body);
      localStorage.clear();
  }

  function getBody(body)
  {
      var res = "";
      body.forEach(function (entry) {
          res += entry + "<br/><br/>";
      })
      return res;
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();
