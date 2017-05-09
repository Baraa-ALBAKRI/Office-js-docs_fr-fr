# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> **Remarque** : dans cette documentation, l’[ensemble de conditions requises](tutorial-api-requirement-sets.html) est présenté en **aperçu**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément. La disponibilité des méthodes et des propriétés présentées dans cet ensemble de conditions doit être testée avant de les utiliser.

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises de la version 1.5](../1.5/index.md). 

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

- [Event.Completed](Event.md#completed) : nouveau paramètre facultatif `options`, qui est un dictionnaire ayant comme seule valeur valide `allowEvent`. Cette valeur est utilisée pour annuler l’exécution d’un événement.

## <a name="additional-resources"></a>Ressources supplémentaires

- [Compléments Outlook](../../../docs/outlook/outlook-add-ins.md)
- [Exemples de code pour les compléments Outlook](https://dev.outlook.com/MailAppsGettingStarted/Samples)
- [Prise en main](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
