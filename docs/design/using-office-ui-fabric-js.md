-
#<a name="use-office-ui-fabric-in-office-add-ins"></a>Utilisation d’Office UI Fabric dans des compléments Office

Si vous créez un complément Office, nous vous encourageons à utiliser [Office UI Fabric](https://dev.office.com/fabric) pour mettre au point l’expérience utilisateur. 

Office UI Fabric est une infrastructure frontale JavaScript permettant de créer des expériences pour Office et Office 365. Fabric propose des composants axés sur des visuels que vous pouvez étendre, retravailler et utiliser dans votre complément Office. Fabric utilisant le langage de création d’Office, ses composants d’expérience utilisateur ressemblent à une extension naturelle d’Office.

La structure se compose de plusieurs projets :

- **Composants JS Fabric (recommandé)** : implémente les composants UX à l’aide d’un code JavaScript uniquement. Nous recommandons d’utiliser cette version de la structure si vous ne souhaitez pas dépendre de l’infrastructure React.  
- **Fabric React** : implémente les composants UX à l’aide de l’infrastructure React.
- **Fabric Core** : contient les principaux éléments du langage de création tels que les icônes, les couleurs, le type et la grille. Les composants JS et Fabric React utilisent les Fabric Core. 

La procédure suivante présente les opérations de base pour l’utilisation de cette structure JS.  

##<a name="1-add-the-fabric-cdn-references"></a>1. Ajouter les références CDN de la structure
Pour faire référence à la structure à partir de CDN, ajoutez le code HTML suivant à votre page.

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

Voilà, vous êtes maintenant prêt à commencer à utiliser la structure dans votre complément. 

##<a name="2-use-fabric-icons-and-fonts"></a>2. Utiliser les polices et les icônes de la structure
Les icônes sont très simples à utiliser. Il vous suffit d’utiliser un élément « i » et de référencer les classes appropriées. Vous pouvez contrôler la taille de l’icône en modifiant la taille de police. Par exemple, le code suivant montre comment créer une icône de tableau extra large qui utilise la couleur themePrimary (#0078 d 7). 
   
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>

Pour rechercher des icônes supplémentaires disponibles dans Office UI Fabric, utilisez la fonctionnalité de recherche de la page [Icônes](https://dev.office.com/fabric#/styles/icons). Lorsque vous trouvez une icône à utiliser dans votre complément, veillez à précéder le nom de l’icône de `ms-Icon--`. 

Pour plus d’informations sur les tailles de police et les couleurs disponibles dans Office UI Fabric, voir [Typographie](https://dev.office.com/fabric#/styles/typography) et [Couleurs](https://dev.office.com/fabric#/styles/colors).

##<a name="3-use-fabric-js-ux-components"></a>3. Utiliser les composants UX de la structure JS

La structure fournit plusieurs composants UX, tels que des boutons ou cases à cocher, que vous pouvez utiliser dans votre complément. Voici une liste des composants UX de la structure JS que nous vous recommandons d’utiliser dans un complément. Pour utiliser l’un des composants de la structure dans votre complément, suivez le lien vers la documentation de la structure, puis suivez les instructions de la section **Utilisation de ce composant**.

> **Remarque :** nous allons ajouter des composants supplémentaires au fil du temps. 

- [Breadcrumb](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Bouton](https://dev.office.com/fabric-js/Components/Button/Button.html) (Envisagez d’utiliser la variante bouton de petite taille dans votre complément. Ajoutez 16 px de marge intérieure aux boutons de petite taille pour garantir une cible tactile de 40 px au minimum sur les appareils tactiles).
- [Checkbox](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Sélecteur de dates](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (pour un exemple de mise en œuvre du sélecteur de dates dans un complément, voir l’exemple de code [Suivi de ventes Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).)
- [Liste déroulante](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Étiquette](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Lien](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [Liste](https://dev.office.com/fabric-js/Components/List/List.html) (vous pouvez modifier les styles par défaut du composant dans le fichier CSS.)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Superposition](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Volet](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Pivot](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Zone de recherche](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Bouton fléché](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Tableau](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Bouton bascule](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Mise à jour de votre complément pour utiliser la structure JS
Si vous utilisez une version précédente d’Office UI Fabric et que vous souhaitez migrer vers Fabric JS, assurez-vous que vous connaissez, incorporez et testez les nouveaux composants de votre complément. Gardez les points suivants à l’esprit pour vous aider à planifier vos mises à jour :

- L’initialisation des composants est plus simple à l’aide de la structure JS. Pour les versions précédentes de la structure, vous incluez le fichier JavaScript du composant de la structure dans votre projet de complément, incluez une référence `<Script>` à ce fichier, puis initialisez le composant. Dans la structure JS, vous n’avez plus besoin d’inclure le fichier JavaScript du composant de la structure et la référence `<Script>` associée. Il vous suffit d’initialiser le composant de la structure.   
- Plusieurs composants fournissent désormais des fonctions qui contrôlent le comportement du composant UX. Par exemple, le contrôle de case à cocher a une fonction `toggle` qui permet de basculer entre les états activé et désactivé. 
- Certains noms de classe d’icône et styles ont été mis à jour.
- La modification la plus notable consiste à utiliser l’élément `<label>` dans de nombreux composants. L’élément `<label>` contrôle le style du composant. Vous devrez peut-être mettre à jour votre code UX pour utiliser l’élément `<label>`. Par exemple, la modification de la valeur de l’attribut coché de l’élément `<input>` sur une case à cocher de la structure JS n’a aucun effet sur celle-ci. À la place, vous utilisez les fonctions `check`, `unCheck` ou `toggle`.   

##<a name="next-steps"></a>Étapes suivantes
Si vous recherchez un exemple de code de bout en bout qui vous montre comment utiliser la structure JS, nous avons tout prévu. Consultez la ressource suivante :

- [Suivi des ventes d’Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

##<a name="related-resources"></a>Ressources connexes
Si vous cherchez des exemples de code ou de la documentation sur une version précédente de la structure, consultez les rubriques suivantes :

- [Modèles de conception de l’expérience utilisateur (utilise la structure 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Exemples d’éléments d’interface utilisateur Fabric pour les compléments Office (utilise Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Utilisation de la structure 2.6.1 dans un complément Office](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

