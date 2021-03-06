# <a name="design-icons-for-add-in-commands"></a>Concevoir des icônes pour les commandes de complément

[Commandes de complément](add-in-commands.md) Ajoutez des boutons, du texte et des icônes à l’interface utilisateur Office. Vos boutons de commande de complément doivent fournir des icônes significatives et des étiquettes qui identifient clairement l’action que l’utilisateur effectue lorsqu’il utilise une commande. Cet article fournit des instructions stylistiques et de production pour vous aider à concevoir des icônes s’intégrant parfaitement avec Office. 

## <a name="office-icon-design-principles"></a>Principes de conception des icônes Office

La version Office 2013 des clients de bureau Office inclut une iconographie actualisée. La modification stylistique de remplacement est une réduction. Les nouvelles icônes incluent uniquement les éléments de communication essentiels. Les éléments non essentiels, tels que la source de lumière, les dégradés et les perspectives, sont supprimés. Les icônes simplifiées prennent en charge l’analyse rapide des commandes et des contrôles. Suivez ce style pour mieux correspondre à Office.

Les icônes Office sont basées sur les principes de conception suivants : 

- Interprétation moderne de la collection d’icônes Office 
- À la fois nouveau et familier  
- Simple, clair et direct 

L’image suivante montre les icônes qui appliquent les principes de conception modernes.

![Image illustrant les anciennes icônes Office et l’interprétation moderne actualisée des icônes](../../images/icons_image.PNG)

## <a name="icon-guidelines"></a>Instructions relatives aux icônes
Suivez ces instructions lorsque vous créez vos icônes : 

- Respectez la grille 1 px et utilisez l’outil d’édition des images bitmap pour de meilleurs résultats.  
- Renouvelez sans redimensionner. Lorsque vous redimensionnez vos icônes à des tailles supérieures ou inférieures, prenez le temps de redessiner les découpages, les coins et des bords arrondis pour optimiser la netteté de ligne. 
- Supprimez les artefacts qui rendent votre icône désordonnée.
- Ne réutilisez pas les icônes d’Office UI Fabric dans le ruban Office ou le menu contextuel. Les icônes de structure sont stylistiquement différentes et ne correspondront pas. 
- Évitez de vous fier à votre logo ou marque pour communiquer ce que fait une commande de complément. Les repères de marque ne sont pas toujours reconnaissables sur des icônes de petites tailles et lorsque des modificateurs sont appliqués. Les repères de marque entrent souvent en conflit avec les styles d’icônes du ruban Office et peuvent gêner l’attention de l’utilisateur dans un environnement saturé.
- Utilisez un remplissage blanc pour améliorer l’accessibilité. La plupart des objets dans les icônes nécessitent un arrière-plan blanc pour être lisibles sur les thèmes de l’interface utilisateur d’Office et en mode contraste élevé.  
- Utilisez le format PNG avec un arrière-plan transparent. 
- Évitez le contenu localisable dans les icônes, y compris les caractères typographiques, les paragraphes en drapeau et les points d’interrogation. 
- Ne réutilisez pas les métaphores visuelles pour différentes commandes. L’utilisation de la même icône pour différentes actions peut semer la confusion. 
- Simplifiez au maximum le nom de vos boutons. Utilisez une combinaison d’informations visuelles et textuelles pour transmettre sa signification. 


## <a name="icon-size-recommendations-and-requirements"></a>Configuration requise et recommandations sur la taille des icônes

Les icônes du bureau Office 2016 sont des images bitmap. Différentes tailles apparaissent en fonction du paramètre PPP de l’utilisateur et du mode tactile. Incluez les huit tailles prises en charge pour créer la meilleure expérience possible dans tous les contextes et résolutions pris en charge. Voici les tailles prises en charge - trois sont obligatoires :

- 16 px (obligatoire)
- 20 px
- 24 px
- 32 px (obligatoire)
- 40 px
- 48 px
- 64 px (recommandé, meilleur choix pour Mac)
- 80 px (obligatoire)  

Veillez à renouveler les icônes pour chaque taille au lieu de les réduire pour les ajuster.

![Illustration présentant la recommandation qui indique de redimensionner les icônes plutôt que de les réduire](../../images/icon_resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

>**Note:** At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Mise en page et structure de l’icône

Les icônes Office sont généralement constituées d’un élément de base avec des modificateurs d’action et conceptuels superposés. Les modificateurs d’action représentent des concepts tels qu’ajouter, ouvrir, nouveau ou fermer. Les modificateurs conceptuels représentent l’état, l’altération ou une description de l’icône. 

Pour créer des commandes qui s’alignent sur l’interface utilisateur d’Office, suivez les instructions de mise en forme pour les éléments de base et les modificateurs. Cela garantit que vos commandes auront un aspect professionnel et que vos clients auront confiance en votre complément. Si vous apportez des exceptions à ces instructions, faites-le intentionnellement.

L’image suivante montre la disposition des éléments de base et modificateurs dans une icône Office.

![Image illustrant un élément de base d’icône dans le centre avec un modificateur dans le coin inférieur droit et un modificateur d’action dans le coin supérieur gauche](../../images/icon_layout.PNG)

- Éléments de base centraux dans le cadre de pixel avec remplissage vide tout autour.
- Placez les modificateurs d’action dans le coin supérieur gauche. 
- Placez les modificateurs conceptuels dans la partie inférieure droite.
- Limitez le nombre d’éléments dans les icônes. En 32 px, limitez le nombre de modificateurs à un maximum de deux. En 16 px, limitez le nombre de modificateurs à un.

Placez les éléments de base de façon cohérente en fonction des tailles. Si les éléments de base ne peuvent pas être centrés dans le cadre, alignez-les en haut à gauche, en laissant les pixels supplémentaires dans la partie inférieure droite. Pour obtenir de meilleurs résultats, appliquez les instructions de remplissage répertoriées dans le tableau suivant.

|**Taille de l’icône**|**Remplissage autour de l’élément de base**|
|:---|:---|
|16px|0|
|20px|1px|
|24px|1px|
|32px|2px|
|40px|2px|
|48px|3px|
|64px|5px|
|80px|5px|

Tous les modificateurs doivent avoir un découpage transparent 1 px entre chaque élément, y compris l’arrière-plan. Les éléments ne doivent pas se chevaucher directement. Créez des espaces entre les règles et les bords. Les modificateurs peuvent varier légèrement en taille, mais utilisez ces dimensions comme point de départ.

|**Taille de l’icône**|**Taille du modificateur**|
|:---|:---|
|16px|9px|
|20px|10px|
|24px|12px|
|32px|14px|
|40px|20px|
|48px|22px|
|64px|29px|
|80px|38px|

## <a name="icon-colors"></a>Couleurs de l’icône

Les icônes Office ont une palette de couleurs limitée. Utilisez les couleurs répertoriées dans le tableau suivant pour garantir une intégration parfaite avec l’interface utilisateur d’Office. Appliquez les instructions suivantes sur l’utilisation des couleurs : 

- Utilisez la couleur pour véhiculer une signification plutôt que pour embellir. Elle doit mettre en surbrillance ou mettre en évidence une action, un état ou un élément qui différencie explicitement le repère.  
- Si possible, n’utilisez qu’une seule couleur supplémentaire au-delà du gris. Limitez les couleurs supplémentaires à deux au maximum.
- Les couleurs ont une apparence cohérente dans toutes les tailles d’icône. Les icônes Office ont des palettes de couleurs légèrement différentes pour des tailles d’icônes différentes. Les icônes 16 px et plus petites sont légèrement plus sombres et plus percutantes que les icônes 32 px et plus grandes. Sans ces ajustements discrets, les couleurs semblent varier en taille.   

|**Nom de la couleur**|**RVB**|**Hex**|**Couleur**|**Catégorie**|
|:---|:---|:---|:---|:---|
|Texte gris (80)|80, 80, 80|#505050|![Image couleur texte gris 80](../../images/textGray_80.gif)|Texte|
|Texte gris (95)|95, 95, 95|#5F5F5F|![Image couleur texte gris 95](../../images/textGray_95.gif)|Texte|
|Texte gris (105)|105, 105, 105|#696969|![Image couleur texte gris 105](../../images/textGray_105.gif)|Texte|
|Gris foncé 32|128, 128, 128|#808080|![Image couleur gris foncé 32](../../images/darkGray_32.gif)|32 et plus|
|Gris moyen 32|158, 158, 158|#9E9E9E|![Image couleur gris moyen 32](../../images/mediumGray_32.gif)|32 et plus|
|TOUT gris clair|179, 179, 179|#B3B3B3|![Image couleur tout en gris clair](../../images/lightGray_all.gif)|Toutes les tailles|
|Gris foncé 16|114, 114, 114|#727272|![Image couleur gris foncé 16](../../images/darkGray_16.gif)|16 et moins|
|Gris moyen 16|144, 144, 144|#909090|![Image couleur gris moyen 16](../../images/mediumGray_16.gif)|16 et moins|
|Bleu 32|77, 130, 184|#4d82B8|![Image couleur bleu 32](../../images/blue_32.gif)|32 et plus|
|Bleu 16|74, 125, 177|#4A7DB1|![Image couleur bleu 16](../../images/blue_16.gif)|16 et moins|
|TOUT jaune|234, 194, 130|#EAC282|![Image couleur tout en jaune](../../images/yellow_all.gif)|Toutes les tailles|
|Orange 32|231, 142, 70|#E78E46|![Image couleur orange 32](../../images/orange_32.gif)|32 et plus|
|Orange 16|227, 142, 70|#E3751C|![Image couleur orange 16](../../images/orange_16.gif)|16 et moins|
|TOUT rose|230, 132, 151|#E68497|![Image couleur tout en rose](../../images/pink_all.gif)|Toutes les tailles|
|Vert 32|118, 167, 151|#76A797|![Image couleur vert 32](../../images/green_32.gif)|32 et plus|
|Vert 16|104, 164, 144|#68A490|![Image couleur 16 vert](../../images/green_16.gif)|16 et moins|
|Rouge 32|216, 99, 68|#D86344|![Image couleur rouge 32](../../images/red_32.gif)|32 et plus|
|Rouge 16|214, 85, 50|#D65532|![Image couleur rouge 16](../../images/red_16.gif)|16 et moins|
|Violet 32|152, 104, 185|#9868B9|![Image couleur violet 32](../../images/purple_32.gif)|32 et plus|
|Violet 16|137, 89, 171|#8959AB|![Image couleur violet 16](../../images/purple_16.gif)|16 et moins|


## <a name="additional-resources"></a>Ressources supplémentaires

- [Meilleures pratiques en matière de développement de compléments](../overview/add-in-development-best-practices.md)
- [Commandes de complément pour Excel, Word et PowerPoint](../design/add-in-commands.md)
