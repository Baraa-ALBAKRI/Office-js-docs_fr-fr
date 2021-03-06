# <a name="onenote-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour OneNote

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Le tableau suivant répertorie les ensembles de conditions requises pour OneNote, les applications hôtes Office qui prennent en charge ces conditions et les numéros de version ou la date de disponibilité.

|  Ensemble de conditions requises  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | Septembre 2016 |  

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office
Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>API JavaScript pour OneNote 1.1 
L’API JavaScript 1.1 pour OneNote est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[API JavaScript pour OneNote](../../docs/onenote/onenote-add-ins-programming-overview.md).

## <a name="runtime-requirement-support-check"></a>Vérification de la prise en charge d’un ensemble de conditions requises à l’exécution

Lors de l’exécution, les compléments peuvent vérifier si un hôte particulier prend en charge un ensemble de conditions requises d’API en procédant comme suit : 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Vérification de la prise en charge d’un ensemble de conditions requises basée sur le manifeste

Utilisez l’élément Conditions requises dans le manifeste du complément pour spécifier des ensembles de conditions requises essentiels ou des membres d’API que votre complément doit utiliser. Si la plateforme ou l’hôte Office ne prend pas en charge les ensembles de conditions requises ou les membres d’API spécifiés dans l’élément Conditions requises, le complément ne s’exécute pas dans cet hôte ou cette plateforme et ne s’affiche pas dans Mes compléments.

Cet exemple de code illustre un complément qui se charge dans toutes les applications hôtes Office qui prennent en charge l’ensemble de conditions requises OneNoteApi, version 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```



## <a name="additional-resources"></a>Ressources supplémentaires

- [Spécification des exigences en matière d’hôtes Office et d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../docs/overview/add-in-manifests.md)
