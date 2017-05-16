# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>Créer un complément Office ASP.NET qui utilise l’authentification unique (aperçu)

Les utilisateurs peuvent se connecter à Office et votre complément Web Office peut tirer parti de cette procédure de connexion pour autoriser les utilisateurs de votre complément et de Microsoft Graph sans obliger les utilisateurs à une deuxième authentification. Pour obtenir une vue d’ensemble, voir [Activer l’authentification unique dans un complément Office](../../docs/develop/sso-in-office-add-ins.md).

Cet article vous guide tout au long du processus d’activation de l’authentification unique (SSO) dans un complément intégré avec ASP.NET, OWIN et la bibliothèque d’authentification Microsoft (MSAL) pour .NET. 

> **Remarque :** Pour un article similaire concernant un complément basé sur Node.js, voir [Créer un complément Office Node.js qui utilise l’authentification unique](../../docs/develop/create-sso-office-add-ins-nodejs.md).

## <a name="prerequisites"></a>Conditions préalables

* Visual Studio 2017 Version 15.3 (26424.2-Preview) ou version ultérieure.

* Office 2016, Version 1704, build 8027.nnnn ou version ultérieure (la version par abonnement Office 365, parfois appelée « Démarrer en un clic »). Vous devrez peut-être participer au programme Office Insider pour obtenir cette version. Pour plus d’informations, voir [Participez au programme Office Insider](https://products.office.com/en-us/office-insider?tab=tab-1).

## <a name="set-up-the-starter-project"></a>Configurer le projet de démarrage

1. Clonez ou téléchargez le référentiel sur [Complément Office ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso). 

1. Ouvrez le dossier **Before** et ouvrez le fichier .sln dans Visual Studio. Il s’agit d’un projet de démarrage. L’interface utilisateur et d’autres aspects du complément qui ne sont pas directement liés à l’authentification unique ou à l’autorisation sont déjà terminés. 

    > Remarque : Il existe également une version finale de l’échantillon dans le même référentiel. Elle est équivalente au complément que vous obtiendriez si vous terminiez les procédures de cet article, sauf que le projet terminé comporte des commentaires de code qui seraient redondants avec le texte de cet article. Pour utiliser la version finale, ouvrez simplement le fichier *.sln et suivez les instructions de cet article, mais ignorez les sections **Code côté client** et **Code côté serveur**.

1. Une fois le projet ouvert, générez-le dans Visual Studio, qui installera les packages répertoriés dans le fichier packages.config. L’opération peut prendre de quelques secondes à plusieurs minutes selon le nombre de packages présents dans le cache de packages de l’ordinateur local.

1. Une fois le projet complètement généré, appuyez sur F5. PowerPoint s’ouvre et un groupe **SSO ASP.NET** se trouve sur le ruban **Accueil**. 

1. Appuyez sur le bouton **Afficher le complément** dans ce groupe pour voir une interface utilisateur du complément dans le volet Office. Le bouton dans le volet Office n’est pas encore raccordé en haut. 
2. Dans Visual Studio, arrêtez le débogueur.

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>Enregistrez le complément avec le point de terminaison Azure AD v2.0

1. Accédez à [https://apps.dev.microsoft.com/?test=build2017](https://apps.dev.microsoft.com/?test=build2017) . 

1. Connectez-vous avec les informations d’identification d’administrateur à votre client Office 365. Par exemple, MonNom@contoso.onmicrosoft.com

1. Cliquez sur **Ajouter une application**.

1. Lorsque vous y êtes invité, utilisez « Office-Add-in-ASPNET-SSO » comme nom d’application et appuyez sur **Créer une application**.

1. Quand la page de configuration de l’application s’ouvre, copiez l’**ID de l’application** et enregistrez-le. Vous l’utiliserez dans une procédure ultérieure. 

    > Remarque : Cet ID est la valeur « audience » lorsque d’autres applications, telles que l’application hôte Office (par exemple, PowerPoint, Word, Excel) recherchent un accès autorisé à l’application. Il s’agit également de l’« ID client » de l’application dès que celle-ci recherche un accès autorisé à Microsoft Graph.

1. Dans la section **Secrets de l’application**, appuyez sur **Générer un nouveau mot de passe**. Une boîte de dialogue contextuelle s’ouvre avec un nouveau mot de passe (également appelé « secret de l’application »). *Copiez le mot de passe immédiatement et enregistrez-le avec l’ID de l’application.* Vous en aurez besoin dans une procédure ultérieure. Ensuite, fermez la boîte de dialogue.

1. Dans la section **Plateformes**, cliquez sur **Ajouter une plateforme**. 

1. Dans la boîte de dialogue qui s’ouvre, sélectionnez **API Web**.

1. Un **URI d’ID d’application** a été généré sous la forme « api://{application ID GUID} ». Remplacez le GUID par « localhost:44355 ». L’ID dans son intégralité doit indiquer `api://localhost:44355`. (La partie domaine du nom d’**étendue**, juste en dessous de l’**URI d’ID d’application** change automatiquement en conséquence. Il doit apparaître comme suit : `api://localhost:44355/access_as_user`.)

1. La section **Applications préalablement autorisées** contient une zone **ID d’application** vide. Entrez l’ID suivant dans la zone (il s’agit de l’ID de Microsoft Office) : `d3590ed6-52b3-4102-aeff-aad2292ab01c`.

1. Ouvrez le menu déroulant **Étendue** à côté de l’**ID d’application** et activez la case à cocher `api://localhost:44355/access_as_user`.

1. En haut de la section **Plateformes**, cliquez sur **Ajouter une plateforme** à nouveau, puis sélectionnez **Web**.

1. Dans la nouvelle section **Web** sous **Plateformes**, entrez les informations suivantes en guise d’**URL de redirection** : `https://localhost:44355`. 

    > Remarque : À ce jour, la plateforme **API Web** disparaît parfois de la section **Plateformes**, tout particulièrement si la page est actualisée après l’ajout de la plateforme **Web** *et l’enregistrement de la page d’inscription*. Pour être sûr que votre plateforme **API Web** fait toujours partie de l’inscription, cliquez sur le bouton **Modifier le manifeste de l’application** près du bas de la page. Vous devriez voir la chaîne `api://localhost:44355` dans la propriété **identifierUris** du manifeste. Il devrait également y avoir une propriété **oauth2Permissions** dont la propriété secondaire **value** a la valeur `access_as_user`.

1. Faites défiler jusqu’à la section **Autorisations pour Microsoft Graph** et à la sous-section **Autorisations déléguées**. Utilisez le bouton **Ajouter** pour ouvrir une boîte de dialogue **Sélectionner des autorisations**.

1. Dans la boîte de dialogue, cochez les cases correspondant aux autorisations suivantes (certaines peuvent être déjà activées par défaut) : 
 * Files.Read.All
 * offline_access
 * openid
 * profil

1. Cliquez sur **OK** au bas de la boîte de dialogue.

1. Cliquez sur **Enregistrer** au bas de la page d’inscription.

## <a name="grant-admin-consent-to-the-add-in"></a>Accorder le consentement de l’administrateur au complément

1. Si le complément ne fonctionne pas dans Visual Studio, appuyez sur F5 pour l’exécuter. Il doit s’exécuter dans IIS pour que cette procédure se déroule sans problème. 

1. Dans la chaîne suivante, remplacez l’espace réservé « {application_ID} » par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément.

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. Collez l’URL résultante dans la barre d’adresses d’un navigateur pour y accéder.

1. Lorsque vous y êtes invité, connectez-vous avec les informations d’identification d’administrateur à votre client Office 365.

1. Vous êtes ensuite invité à accorder des autorisations pour votre complément pour accéder à vos données Microsoft Graph. Cliquez sur **Accepter**. 

1. L’onglet ou la fenêtre du navigateur est alors redirigé vers l’**URL de redirection** que vous avez spécifiée lorsque vous avez enregistré le complément, afin que la page d’accueil du complément s’ouvre dans le navigateur. 

2. Dans la barre d’adresses du navigateur, vous verrez un paramètre de requête « client » avec une valeur GUID. Il s’agit de l’ID de votre client Office 365. Copiez et enregistrez cette valeur. Vous l’utiliserez dans une étape ultérieure.

3. Fermez la fenêtre/l’onglet.

1. Arrêtez le débogueur dans Visual Studio.

## <a name="configure-the-add-in"></a>Configurer le complément

1. Dans la chaîne suivante, remplacez l’espace réservé « {tenant_ID} » par l’ID de client Office 365 que vous avez obtenu précédemment. Si pour une raison quelconque, vous n’avez pas obtenu l’ID antérieur, utilisez l’une des méthodes de la page [Trouver votre ID de client Office 365](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) pour l’obtenir.

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. Dans Visual Studio, ouvrez le fichier web.config. Il existe certaines clés dans la section **appSettings** à laquelle vous devez affecter les valeurs.

1. Utilisez la chaîne que vous avez créée à l’étape 1 en tant que valeur pour la clé nommée « ida:Issuer ». Assurez-vous que la valeur ne comporte aucun espace vide.

1. Donnez les valeurs suivantes aux clés correspondantes :

|Clé|Valeur|
|:-----|:-----|
|ida:ClientID|L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.|
|ida:Audience|L’ID d’application que vous avez obtenu lorsque vous avez enregistré le complément.|
|ida:Password|Le mot de passe que vous avez obtenu lorsque vous avez enregistré le complément.|


Voici un exemple de ce à quoi doivent ressembler les quatre clés que vous avez modifiées. *Notez que les clés ClientID et Audience sont identiques*.

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

> **Remarque :** Conservez les autres paramètres de la section **appSettings** inchangés.


1. Enregistrez et fermez le fichier.

1. Dans le projet de complément, ouvrez le fichier manifeste du complément « Office-Add-in-ASPNET-SSO.xml ».

1. Faites défiler vers le bas du fichier.

1. Juste au-dessus de la balise de fin </VersionOverrides>, vous trouverez le balisage suivant :

    ```xml
    <WebApplicationId>{application_GUID here}</WebApplicationId>
    <WebApplicationResource>api://localhost:44355<WebApplicationResource>
    <WebApplicationScopes>
        <WebApplicationScope>profile</WebApplicationScope>
        <WebApplicationScope>openid</WebApplicationScope>
        <WebApplicationScope>offline_access</WebApplicationScope>
        <WebApplicationScope>files.read.all</WebApplicationScope>
    </WebApplicationScopes>
    ```

1. Remplacez l’espace réservé « {application_GUID} » dans le balisage par l’ID d’application que vous avez copié lorsque vous avez enregistré votre complément. C’est le même ID que celui que vous avez utilisé pour ClientID et Audience dans le fichier web.config.

    >Remarque : 
    >
    >* La valeur **WebApplicationResource** correspond à l’**URI d’ID d’application** défini lorsque vous avez ajouté la plateforme API Web à l’enregistrement du complément.
    >* La section **WebApplicationScopes** est utilisée uniquement pour générer une boîte de dialogue de consentement si le complément est vendu via Office Store.

1. Enregistrez et fermez le fichier.

## <a name="code-the-client-side"></a>Code côté client

1. Ouvrez le fichier Home.js dans le dossier **Scripts**. Il contient déjà du code :

    * Une affectation à la méthode `Office.initialize` qui affecte elle-même un gestionnaire à l’événement ClickButton `getGraphAccessTokenButton`.
    * Une méthode `showResult` permettant d’afficher les données renvoyées par Microsoft Graph (ou un message d’erreur) en bas du volet Office.

1. En dessous de l’affectation au `Office.initialize`, ajoutez le code ci-dessous. Tenez compte des informations suivantes : 

    * `getAccessTokenAsync` est la nouvelle API d’Office.js qui permet à un complément de demander à l’application hôte Office (Excel, PowerPoint, Word, etc.) un jeton d’accès au complément (pour l’utilisateur connecté à Office). L’application hôte Office demande alors le jeton au point de terminaison Azure AD 2. Dans la mesure où vous avez préalablement autorisé l’hôte Office sur votre complément lors de son inscription, Azure AD enverra le jeton. 
    * Si aucun utilisateur n’est connecté à Office, l’hôte Office invite l’utilisateur à se connecter. 
    * Le paramètre options définit `forceConsent` sur false, donc l’utilisateur ne sera pas invité à accorder l’accès de l’hôte Office à votre complément.

    ```js
    function getOneDriveItems() {
    Office.context.auth.getAccessTokenAsync({ forceConsent: false },
        function (result) {
            if (result.status === "succeeded") {
                // TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                console.log("Code: " + result.error.code);
                console.log("Message: " + result.error.message);
                console.log("name: " + result.error.name);
                document.getElementById("getGraphAccessTokenButton").disabled = true;
            }
        });
    }
    ```

1. Remplacez TODO1 par les lignes suivantes. Vous créez la méthode `getData` et la route « /api/values » côté serveur dans les étapes suivantes. Une URL relative est utilisée pour le point de terminaison car il doit être hébergé sur le même domaine que votre complément.

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. En dessous de la méthode `getOneDriveFiles`, ajoutez le code suivant. Cette méthode utilitaire appelle un point de terminaison API Web spécifié et lui transmet le jeton d’accès que l’application hôte Office a utilisé pour accéder à votre complément. Sur le côté serveur, ce jeton d’accès est utilisé dans le flux « de la part de » pour obtenir un jeton d’accès à Microsoft Graph. 

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            console.log(result.error);
        });
    }
    ```

1. Enregistrez et fermez le fichier.

## <a name="code-the-server-side"></a>Code côté serveur

### <a name="configure-the-owin-middleware"></a>Configurer les intergiciels OWIN

1. Ouvrez le fichier Startup.cs à la racine du projet. 

1. Ajoutez le mot clé `partial` à la déclaration de la classe de démarrage, si ce n’est pas déjà fait. Elle doit ressembler à ceci :

    `public partial class Startup`

1. Ajoutez la ligne suivante dans le corps de la méthode `Configure`. Vous créez la méthode `ConfigureAuth` dans une étape ultérieure.

    `ConfigureAuth(app);`

1. Enregistrez et fermez le fichier.

1. Cliquez avec le bouton droit de la souris sur le dossier **App_Start**, puis sélectionnez **Ajouter | Classe**. 

1. Dans la boîte de dialogue **Ajouter un nouvel élément** nommez le fichier **Startup.Auth.cs**, puis cliquez sur **Ajouter**.

1. Raccourcissez le nom de l’espace de noms dans le nouveau fichier `Office_Add_in_ASPNET_SSO_WebAPI`.

1. Vérifiez que toutes les instructions `using` suivantes se trouvent en haut du fichier. 

   ```
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. Ajoutez le mot clé `partial` à la déclaration de la classe `Startup`, si ce n’est pas déjà fait. Elle doit ressembler à ceci :

    `public partial class Startup`

1. Ajoutez la méthode suivante à la classe `Startup`. Cette méthode spécifie comment l’intergiciel OWIN valide les jetons d’accès qui lui sont transmis à partir de la méthode `getData` dans le fichier Home.js côté client. Le processus d’autorisation est déclenché chaque fois qu’un point de terminaison Web API décoré avec l’attribut `[Authorize]` est appelé.

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO2: Configure the validation settings
        // TODO3: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. Remplacez TODO2 par les lignes suivantes. Remarque :

    * Le code demande à OWIN de s’assurer que l’audience et l’émetteur du jeton spécifiés dans le jeton d’accès qui provient de l’hôte Office (et est transmis par l’appel côté client de `getData`) doivent correspondre aux valeurs spécifiées dans le fichier web.config.
    * Le réglage de `SaveSigninToken` sur `true` fait qu’OWIN enregistre le jeton brut à partir de l’hôte Office. Le complément en a besoin pour obtenir un jeton d’accès à Microsoft Graph avec le flux « de la part de ».
    * Les étendues ne sont pas validées par l’intergiciel OWIN. Les étendues du jeton d’accès, qui doivent inclure `access_as_user`, sont validées dans le contrôleur.

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. Remplacez TODO3 par les lignes suivantes. Remarque :

    * La méthode `UseOAuthBearerAuthentication` est appelée au lieu de la méthode `UseWindowsAzureActiveDirectoryBearerAuthentication` plus courante car cette dernière n’est pas compatible avec le point de terminaison Azure AD V2.
    * L’URL de découverte transmise à la méthode correspond à l’endroit où l’intergiciel OWIN obtient les instructions permettant d’obtenir la clé requise pour vérifier la signature sur le jeton d’accès reçu de l’hôte Office.

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. Enregistrez et fermez le fichier.

### <a name="create-the-apivalues-controller"></a>Créer le contrôleur /api/values

1. Ouvrez le fichier **Controllers\ValueController.cs**. 

1. Vérifiez que les instructions `using` suivantes se trouvent en haut du fichier.

    ```
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

1. Juste au-dessus de la ligne qui déclare le `ValuesController`, ajoutez l’attribut `[Authorize]`. Cela permet de s’assurer que votre complément exécutera le processus d’autorisation que vous avez configuré dans la dernière procédure chaque fois qu’une méthode de contrôleur est appelée ; seuls les appelants avec un jeton d’accès valide à votre complément peuvent ainsi appeler les méthodes du contrôleur. 

1. Ajoutez la méthode suivante au `ValuesController` :

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO4: Validate the scopes of the access token.
    }
    ```

1. Remplacez TODO4 par le code suivant pour confirmer que les étendues qui sont spécifiées dans le jeton incluent `access_as_user`. 

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO5: Get the raw token that the add-in page received from the Office host.
        // TODO6: Get the access token for MS Graph.
        // TODO7: Get the names of files and folders in OneDrive for Business by using the Microsoft Graph API.
        // TODO8: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

1. Remplacez TODO5 par le code suivant qui transforme le jeton d’accès brut reçu de l’hôte Office en objet `UserAssertion` qui sera transmis à une autre méthode.

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ```

1. Remplacez TODO6 par le code suivant. Remarque :

    * Votre complément ne joue plus le rôle d’une ressource (ou audience) à laquelle l’hôte Office et l’utilisateur doivent accéder. Désormais, il est lui-même un client qui a besoin d’accéder à Microsoft Graph. `ConfidentialClientApplication` est l’objet de « contexte client » MSAL. 
    * Le troisième paramètre du constructeur `ConfidentialClientApplication` est une URL de redirection qui n’est pas utilisée dans le flux « de la part de », mais il est recommandé d’utiliser l’URL correcte. Les quatrième et cinquième paramètres peuvent être utilisés pour définir un magasin permanent qui permettrait la réutilisation des jetons non expirés entre différentes sessions avec le complément. Cet exemple n’implémente pas un stockage permanent.
    * La méthode `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` recherchera tout d’abord dans le cache MSAL, c’est-à-dire en mémoire, un jeton d’accès correspondant. Uniquement s’il n’existe pas, elle lance le flux « de la part de » avec le point de terminaison Azure AD V2.

    ```
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "profile", "Files.Read.All" };
    AuthenticationResult result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    ```

1. Remplacez TODO7 par les lignes suivantes. Remarque :

    * Les classes `GraphApiHelper` et `ODataHelper` sont définies dans les fichiers du dossier **Helpers**. La classe `OneDriveItem` est définie dans un fichier du dossier **Models**. La description détaillée de ces classes n’est pas pertinente pour l’autorisation ou l’authentification unique, elle est donc hors de portée de cet article.
    * Vous pouvez améliorer les performances en demandant à Microsoft Graph uniquement les données réellement requises, pour le code utilise un paramètre de requête ` $select` pour spécifier que nous ne souhaitons que la propriété name et un paramètre `$top` pour spécifier que nous ne voulons que les 3 premiers dossiers de noms de fichiers.

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

1. Remplacez TODO8 par les lignes suivantes. Notez que bien que le code ci-dessus demande uniquement la propriété *name* des éléments OneDrive, Microsoft Graph comporte toujours la propriété *eTag* pour les éléments OneDrive. Pour réduire la charge utile envoyée au client, le code ci-dessous reconstruit les résultats avec uniquement les noms d’élément.

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## <a name="run-the-add-in"></a>Exécution du complément

1. Vérifiez que vous avez des fichiers dans votre espace OneDrive Entreprise.

1. Dans Visual Studio, appuyez sur F5. PowerPoint s’ouvre et un groupe **SSO ASP.NET** se trouve sur le ruban **Accueil**. 

1. Appuyez sur le bouton **Afficher le complément** dans ce groupe pour voir une interface utilisateur du complément dans le volet Office. 

1. Appuyez sur le bouton **Obtenir mes fichiers à partir de** OneDrive. Si vous n’êtes pas connecté à Office, vous êtes invité à vous connecter.

1. Une fois que vous êtes connecté, une liste de vos fichiers et dossiers dans OneDrive Entreprise s’affiche sous le bouton. Cette opération peut prendre plus de 15 secondes, surtout la première fois. 



