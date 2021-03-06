
# <a name="host-element"></a>Élément Host
Spécifie un type d’application Office individuel dans lequel le complément doit s’activer.

> **Important** : La syntaxe des éléments **Host** varie selon que l’élément est défini dans le [manifeste de base](#basic-manifest) ou le nœud [VersionOverrides](#versionoverrides-node). Toutefois, la fonctionnalité est identique.  


## <a name="basic-manifest"></a>Manifeste de base

Lorsqu’il est défini dans le manifeste base (sous [OfficeApp](./officeapp.md)), le type d’hôte est déterminé par l’attribut `Name`.   

### <a name="attributes"></a>Attributs
| Attribut     | Type   | Requis | Description                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [Name](#name) | chaîne | obligatoire | Nom du type d’application hôte Office. |


### <a name="name"></a>Nom
Spécifie le type d’hôte ciblé par ce complément. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `Database` (Access)
- `Mailbox` (Outlook)
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Project` (Project)
- `Workbook` (Excel)

### <a name="example"></a>Exemple
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

---

## <a name="versionoverrides-node"></a>Nœud VersionOverrides
Lorsqu’il est défini dans [VersionOverrides](./versionoverrides.md), le type d’hôte est déterminé par l’attribut `xsi:type`. 

### <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [xsi:type](#xsitype)  |  Oui  | Décrit l’hôte d’Office auquel ces paramètres s’appliquent.|

### <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  [DesktopFormFactor](./desktopformfactor.md)    |  Oui   |  Définit les paramètres pour le facteur de forme pour bureau. |
|  [MobileFormFactor](./mobileformfactor.md)    |  Non   |  Définit les paramètres pour le facteur de forme pour environnement mobile. **Remarque :** cet élément est uniquement pris en charge dans Outlook pour iOS. |


### <a name="xsitype"></a>xsi:type
Contrôle à quel hôte Office (Word, Excel, PowerPoint, Outlook, OneNote) s’appliquent également les paramètres contenus. La valeur doit être l’une des suivantes :

- `Document` (Word)
- `MailHost` (Outlook)    
- `Notebook` (OneNote)
- `Presentation` (PowerPoint)
- `Workbook` (Excel)

## <a name="host-example"></a>Exemple d’hôte 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
