
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            initVuejs();
        });
    };

    var app;
    function initVuejs() {
        // On initialise le modèle de notre page
        // Rien n'empêche de faire une requête Ajax, on est dans un vrai site web
        var dossiers = [
            { id: 1, immat: "GL-984-LG", nomAssure: "M. Guillaume Lacasa", valeur: "1200" },
            { id: 2, immat: "IM-123-MI", nomAssure: "M. Ionut Mihalcea", valeur: "1200" },
            { id: 3, immat: "PL-987-LP", nomAssure: "M. Patrice Lamarche", valeur: "1200" },
            { id: 4, immat: "DT-275-TD", nomAssure: "M. David Toussaint", valeur: "1200" },
        ];
        app = new Vue({
            el: '#content-main',
            data: {
                dossiers: dossiers,
                selection: null,
                insert: insert
            },
        });
    }

    function insert(prop) {
        // Word.run permet d'exécuter le code dans Word.
        // Ici on s'en sert pour ajouter du texte dans le document
        Word.run(function (context) {
            var selection = context.document.getSelection();

            selection.insertText(
                app.selection[prop].toString(),
                Word.InsertLocation.replace);

            // On termine toujours avec context.sync() pour rendre la main à Word
            return context.sync();
        }).catch(errorHandler);
    }


    // Ci-dessous : le code du template de base pour gérer les erreur et afficher des notifications

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
