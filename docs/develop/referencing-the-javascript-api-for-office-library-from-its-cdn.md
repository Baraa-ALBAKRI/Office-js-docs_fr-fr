
# <a name="referencing-the-javascript-api-for-office-library-from-its-content-delivery-network-cdn"></a>Référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu


La bibliothèque de l’[interface API JavaScript pour Office](../../reference/javascript-api-for-office.md) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. 


La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

La valeur `/1/` devant `office.js` dans l’URL CDN préconise l’utilisation de la dernière version incrémentielle comprise dans la version 1 d’Office.js. Étant donné que l’interface API JavaScript pour Office maintient la compatibilité descendante, la dernière version continuera de prendre en charge les membres de l’API ajoutés précédemment dans la version 1. Si vous devez mettre à jour un projet existant, consultez la rubrique relative à la [mise à jour de la version de votre interface API JavaScript pour Office et des fichiers de schéma de manifeste](../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si vous envisagez de publier votre complément Office à partir de l’Office Store, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.

> **Important :** Quand vous développez un complément pour une application hôte Office, veillez à référencer l’interface API JavaScript pour Office depuis l’intérieur de la section `<head>` de la page. Ainsi, l’API est entièrement initialisée avant les éléments Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Au-delà de ce délai, un message d’erreur indiquant que le complément ne répond pas s’affiche à l’écran.       

## <a name="additional-resources"></a>Ressources supplémentaires



- [Présentation de l’API JavaScript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Vue d’ensemble de la plateforme des compléments pour Office](../../docs/overview/office-add-ins.md)
    
- [Cycle de vie du développement des compléments Office](../../docs/design/add-in-development-lifecycle.md)
    
- [Interface API JavaScript pour Office](../../reference/javascript-api-for-office.md)
    
