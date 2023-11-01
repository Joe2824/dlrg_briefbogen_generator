# DLRG Briefbogen Generator

Der DLRG Briefbogen Generator ist eine Python-Anwendung, die individualisierte Briefbögen für die DLRG (Deutsche Lebens-Rettungs-Gesellschaft) generiert. Die Anwendung ermöglicht es Benutzern, vordefinierte Microsoft Word-Vorlagen mit Daten aus Excel-Dateien zu befüllen und individuelle Briefbögen für verschiedene Mitglieder der Organisation zu erstellen.
        
## Download
[Briefbogen Generator](https://github.com/joe2824/dlrg_briefbogen_generator/releases/latest/download/DLRG_Briefbogen_Generator.zip) |
[Briefbogen Templates](https://github.com/joe2824/dlrg_briefbogen_generator/releases/latest/download/Briefbogen_Templates.zip)


## Verwendung

### Template-Ordner auswählen: 
Klicken Sie auf die Schaltfläche "Auswählen" neben "Template Ordner", um das Verzeichnis auszuwählen, das Ihre Microsoft Word-Vorlagen (.docx-Dateien) enthält.

### Ausgabe-Ordner auswählen: 
Klicken Sie auf die Schaltfläche "Auswählen" neben "Ausgabe Pfad", um das Verzeichnis festzulegen, in dem die generierten Briefbögen gespeichert werden.

### Datenquelle auswählen: 
Klicken Sie auf die Schaltfläche "Auswählen" neben "Daten Quelle", um eine Excel-Datei auszuwählen, die die Daten enthält, die in die Vorlagen eingefügt werden sollen. Die Daten-Datei sollte spezifische Tabellenblätter namens "Allgemeine Daten", "Vorstand" und "Jugend" haben.

### Daten-Gruppe auswählen: 
Wählen Sie aus, ob Sie Briefbögen für die "Vorstand" oder die "Jugend" generieren möchten, indem Sie auf das entsprechende Optionsfeld klicken.

### Vorlagen auswählen:
Markieren Sie die Vorlagen, für die Sie Briefbögen auswählen möchten, aus der Liste der Vorlagen im scrollbaren Bereich.

### Briefbögen generieren: 
Klicken Sie auf die Schaltfläche "Generiere Briefbogen", um den Generierungsprozess zu starten. Das Programm wird die Daten aus der ausgewählten Gruppe (Vorstand oder Jugend) mit den ausgewählten Vorlagen zusammenführen und individuelle Briefbögen für jedes Mitglied erstellen.

### Datenbeispiel exportieren: 
Um eine Beispieldaten-Datei zu erstellen, klicken Sie auf die Schaltfläche "Daten Vorlage exportieren" in der Symbolleiste. Dadurch wird eine Excel-Datei mit Beispieldaten generiert, die Sie als Vorlage verwenden können, um Ihre spezifischen Details einzufügen.


## Unterstützte Variablen

Die folgenden Variablen können in den Vorlagen verwendet werden. Variablen, die mit "o_" beginnen, sind spezifisch für die Organisation, während solche, die mit "jo_" beginnen, für die Jugendorganisation sind.

## Hinweis zu den Variablen
RichText Variablen müssen in der Word Vorlage mit einem `r` angeführt werden z.B. `{{r o_email }}`.

Weitere Infos gibt es in der [python-docx-template Dokumentation](https://docxtpl.readthedocs.io/en/latest/)

### Variablen für die Organisation:
- `o_strasse` : Straßenadresse der Organisation
- `o_plz_ort` : Postleitzahl und Stadt der Organisation
- `o_telefon` : Telefonnummer der Organisation
- `o_fax` : Faxnummer der Organisation
- `o_email` : E-Mail-Adresse der Organisation (anklickbarer Link)
- `o_webseite` : Website der Organisation (anklickbarer Link)
- `o_amtsgericht_ort_nummer` : Amtsgericht und Registriernummer der Organisation
- `o_vorsitz` : Name des Vorsitzenden der Organisation
- `o_stv_vorsitz` : Name des stellvertretenden Vorsitzenden der Organisation
- `o_schatzmeister` : Name des Schatzmeisters der Organisation
- `o_bank_1_name` : Name der ersten Bank der Organisation
- `o_bank_1_iban` : IBAN der ersten Bank der Organisation
- `o_bank_1_bic` : BIC der ersten Bank der Organisation
- `o_bank_2_name` : Name der zweiten Bank der Organisation
- `o_bank_2_iban` : IBAN der zweiten Bank der Organisation
- `o_bank_2_bic` : BIC der zweiten Bank der Organisation
- `o_ust_o_str` : Umsatzsteuer-Identifikationsnummer der Organisation
- `o_var1` bis `o_var5` : Benutzerdefinierte Variablen für die Organisation

### Variablen für die Jugendorganisation:
- `jo_strasse` : Straßenadresse der Jugendorganisation
- `jo_plz_ort` : Postleitzahl und Stadt der Jugendorganisation
- `jo_telefon` : Telefonnummer der Jugendorganisation
- `jo_fax` : Faxnummer der Jugendorganisation
- `jo_email` : E-Mail-Adresse der Jugendorganisation (anklickbarer Link)
- `jo_webseite` : Website der Jugendorganisation (anklickbarer Link)
- `jo_amtsgericht_ort_nummer` : Amtsgericht und Registriernummer der Jugendorganisation
- `jo_vorsitz` : Name des Vorsitzenden der Jugendorganisation
- `jo_stv_vorsitz` : Name des stellvertretenden Vorsitzenden der Jugendorganisation
- `jo_schatzmeister` : Name des Schatzmeisters der Jugendorganisation
- `jo_bank_1_name` : Name der ersten Bank der Jugendorganisation
- `jo_bank_1_iban` : IBAN der ersten Bank der Jugendorganisation
- `jo_bank_1_bic` : BIC der ersten Bank der Jugendorganisation
- `jo_bank_2_name` : Name der zweiten Bank der Jugendorganisation
- `jo_bank_2_iban` : IBAN der zweiten Bank der Jugendorganisation
- `jo_bank_2_bic` : BIC der zweiten Bank der Jugendorganisation
- `jo_ust_o_str` : Umsatzsteuer-Identifikationsnummer der Jugendorganisation
- `jo_kreisjugendring` : Name des Jugendkreisrates der Jugendorganisation
- `jo_var1` bis `jo_var5` : Benutzerdefinierte Variablen für die Jugendorganisation