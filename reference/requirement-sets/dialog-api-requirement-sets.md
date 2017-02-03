
# <a name="dialog-api-requirement-sets"></a>Ensembles de conditions requises de l’API de boîte de dialogue

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément. Pour plus d’informations, consultez la rubrique [Spécifier les hôtes Office et les conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Les compléments Office s’exécutent sur plusieurs versions d’Office. Le tableau suivant répertorie les ensembles de conditions requises de l’API de boîte de dialogue, les applications Office hôte qui prennent en charge ces conditions et les numéros de build ou de version de l’application Office.

|  Ensemble de conditions requises  |  Office 2013 pour Windows | Office 2016 pour Windows*   |  Office 2016 pour iPad  |  Office 2016 pour Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | Build 15.0.4855.1000 ou version ultérieure | Version 1602 (Build 6741.0000) ou version ultérieure | 1.22 ou version ultérieure | 15.20 ou version ultérieure| Janvier 2017 | Version 1608 (Build 7601.6800) ou version ultérieure|

>**Remarque :** Le numéro de build d’Office 2016 installé via MSI est 16.0.4266.1001. Pour utiliser l’Api de boîte de dialogue, effectuez la mise à jour d’Office pour obtenir la dernière version. 

Pour en savoir plus sur les versions, les numéros de build et Office Online Server, voir :

- 
  [Numéros de version et de build des canaux de réception des mises à jour pour les clients Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [Quelle est la version d’Office que j’utilise ?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- 
  [Où trouver le numéro de version et de build pour une application cliente Office 365](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- 
  [Présentation d’Office Online Server](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## <a name="office-common-api-requirement-sets"></a>Ensembles de conditions requises des API communes pour Office
Pour plus d’informations sur les ensembles de conditions requises des API communes, voir [Ensembles de conditions requises des API communes pour Office](office-add-in-requirement-sets.md).

## <a name="dialog-api-11"></a>API de boîte de dialogue 1.1 
L’API de boîte de dialogue 1.1 est la première version de l’API. Pour plus d’informations sur l’API, consultez les rubriques de référence sur l’[API de boîte de dialogue](../shared/officeui.md).

## <a name="additional-resources"></a>Ressources supplémentaires

- [Spécification des exigences en matière d’hôtes Office et d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Manifeste XML des compléments Office](../../docs/overview/add-in-manifests.md)

