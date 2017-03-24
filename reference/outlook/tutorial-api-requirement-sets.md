 

# <a name="understanding-outlook-api-requirement-sets"></a>Présentation de l’ensemble de conditions requises pour les API Outlook

Les versions API requises pour les compléments Outlook sont indiquées à l’aide de l’élément [Requirements](https://msdn.microsoft.com/EN-US/library/office/dn592036.aspx) dans leur [manifeste](https://msdn.microsoft.com/en-us/library/office/fp123693.aspx). Les compléments Outlook contiennent toujours un élément [Set](https://msdn.microsoft.com/EN-US/library/office/dn592049.aspx) avec un attribut `Name` défini sur `Mailbox` et un attribut `MinVersion` défini sur l’ensemble minimal de conditions requises de l’API qui prend en charge les scénarios du complément.

Par exemple, l’extrait de manifeste suivant indique l’ensemble minimal de conditions requises 1.1 :

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Toutes les API Outlook appartiennent à l’`Mailbox`[ensemble de conditions requises](https://msdn.microsoft.com/EN-US/library/office/dn535871.aspx#SpecifyRequirementSets_intro). L’ensemble de conditions requises `Mailbox` possède plusieurs versions et chaque nouvel ensemble d’API publié appartient à une version supérieure de l’ensemble. L’ensemble d’API le plus récent n’est pas pris en charge par tous les clients Outlook, mais si un client Outlook prend en charge un ensemble de conditions requises, toutes les API comprises dans cet ensemble sont également prises en charge.

La définition d’une version minimale d’ensemble de conditions requises dans le manifeste permet de contrôler les clients Outlook dans lesquels le complément va apparaître. Si un client ne prend pas en charge l’ensemble minimal de conditions requises, il ne charge pas le complément. Par exemple, si la version de l’ensemble de conditions requises spécifiée est 1.3, le complément n’apparaîtra pas dans les clients Outlook qui ne prennent pas en charge au minimum la version 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Utilisation des API d’un ensemble de conditions requises ultérieure

La définition d’un ensemble de conditions requises ne limite pas votre complément à utiliser les API de cette version. Par exemple, si le complément indique l’ensemble de conditions requises 1.1, mais qu’il s’est exécuté dans un client Outlook prenant en charge la version 1.3, le complément peut utiliser les API de l’ensemble de conditions requises 1.3\.

Pour utiliser des API plus récentes, les développeurs peuvent simplement vérifier leur disponibilité en utilisant la technique JavaScript standard.

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Ces vérifications ne sont pas nécessaires pour les API présentes dans l’ensemble de conditions requises dont la version est la même que celle spécifiée dans le manifeste.

## <a name="choosing-a-minimum-requirement-set"></a>Choix d’un ensemble minimal de conditions requises

Les développeurs doivent utiliser l’ensemble de conditions requises le plus ancien qui contient l’ensemble d’API critique pour leur scénario, sans lequel le complément ne fonctionne pas.

## <a name="clients"></a>Clients

Les clients suivants prennent en charge des compléments Outlook.

| Client | Ensembles de conditions requises des API prises en charge |
| --- | --- |
| Outlook 2016 pour Windows | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2016 pour Mac | 1.1, 1.2, 1.3, 1.4 |
| Outlook 2013 pour Windows | 1.1, 1.2, 1.3 |
| Outlook sur le web (Office 365 et Outlook.com) | 1.1, 1.2, 1.3, 1.4 |
| Outlook Web App (Exchange 2013 sur site) | 1.1 |
| Outlook Web App (Exchange 2016 sur site) | 1.1, 1.2. 1.3 |
>**Remarque** La prise en charge de la version 1.3 dans Outlook 2013 a été ajoutée dans le cadre de la [Mise à jour du 8 décembre 2015 pour Outlook 2013 (KB3114349)](https://support.microsoft.com/en-us/kb/3114349)    
