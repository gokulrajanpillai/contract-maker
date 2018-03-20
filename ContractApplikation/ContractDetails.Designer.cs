namespace ContractApplikation
{
    partial class ContractDetails
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.ContractGenerator = new System.Windows.Forms.TabControl();
            this.Ansprechpartner = new System.Windows.Forms.TabPage();
            this.createCustomerBtn = new System.Windows.Forms.Button();
            this.ort = new System.Windows.Forms.TextBox();
            this.ortLbl = new System.Windows.Forms.Label();
            this.strasse = new System.Windows.Forms.TextBox();
            this.strasseLbl = new System.Windows.Forms.Label();
            this.abteilungszusatz = new System.Windows.Forms.TextBox();
            this.departmentAdditionalLbl = new System.Windows.Forms.Label();
            this.geschäftsbereich = new System.Windows.Forms.TextBox();
            this.departmentLabel = new System.Windows.Forms.Label();
            this.firma = new System.Windows.Forms.TextBox();
            this.firmaLbl = new System.Windows.Forms.Label();
            this.abteilung = new System.Windows.Forms.TextBox();
            this.abteilungLbl = new System.Windows.Forms.Label();
            this.email = new System.Windows.Forms.TextBox();
            this.emailLbl = new System.Windows.Forms.Label();
            this.telefon = new System.Windows.Forms.TextBox();
            this.telefonLbl = new System.Windows.Forms.Label();
            this.nachname = new System.Windows.Forms.TextBox();
            this.nachnameLbl = new System.Windows.Forms.Label();
            this.vorname = new System.Windows.Forms.TextBox();
            this.vornameLbl = new System.Windows.Forms.Label();
            this.frauRadioBtn = new System.Windows.Forms.RadioButton();
            this.herrRadioBtn = new System.Windows.Forms.RadioButton();
            this.Projekt = new System.Windows.Forms.TabPage();
            this.createProjectBtn = new System.Windows.Forms.Button();
            this.ansprechpartnerComboBox = new System.Windows.Forms.ComboBox();
            this.projektBeschreibung = new System.Windows.Forms.TextBox();
            this.disponent = new System.Windows.Forms.TextBox();
            this.gesprächsperson = new System.Windows.Forms.TextBox();
            this.verrechnungssatz = new System.Windows.Forms.TextBox();
            this.anzahlStunden = new System.Windows.Forms.TextBox();
            this.endDatumDtPikr = new System.Windows.Forms.DateTimePicker();
            this.startDatumDtPikr = new System.Windows.Forms.DateTimePicker();
            this.projektBeschreibungLbl = new System.Windows.Forms.Label();
            this.disponentLbl = new System.Windows.Forms.Label();
            this.gesprächspersonLbl = new System.Windows.Forms.Label();
            this.verrechnungssatzLbl = new System.Windows.Forms.Label();
            this.anzahlStundenLbl = new System.Windows.Forms.Label();
            this.ansprechPartnerLbl = new System.Windows.Forms.Label();
            this.endDatumLbl = new System.Windows.Forms.Label();
            this.startDatumLbl = new System.Windows.Forms.Label();
            this.projektnummer = new System.Windows.Forms.TextBox();
            this.projektnummerLbl = new System.Windows.Forms.Label();
            this.Angebot = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.projektComboBox = new System.Windows.Forms.ComboBox();
            this.backgrdDBWorker = new System.ComponentModel.BackgroundWorker();
            this.projektTitel = new System.Windows.Forms.TextBox();
            this.projektTitelLbl = new System.Windows.Forms.Label();
            this.ContractGenerator.SuspendLayout();
            this.Ansprechpartner.SuspendLayout();
            this.Projekt.SuspendLayout();
            this.Angebot.SuspendLayout();
            this.SuspendLayout();
            // 
            // ContractGenerator
            // 
            this.ContractGenerator.Controls.Add(this.Ansprechpartner);
            this.ContractGenerator.Controls.Add(this.Projekt);
            this.ContractGenerator.Controls.Add(this.Angebot);
            this.ContractGenerator.Location = new System.Drawing.Point(12, 12);
            this.ContractGenerator.Name = "ContractGenerator";
            this.ContractGenerator.SelectedIndex = 0;
            this.ContractGenerator.Size = new System.Drawing.Size(955, 523);
            this.ContractGenerator.TabIndex = 0;
            // 
            // Ansprechpartner
            // 
            this.Ansprechpartner.Controls.Add(this.createCustomerBtn);
            this.Ansprechpartner.Controls.Add(this.ort);
            this.Ansprechpartner.Controls.Add(this.ortLbl);
            this.Ansprechpartner.Controls.Add(this.strasse);
            this.Ansprechpartner.Controls.Add(this.strasseLbl);
            this.Ansprechpartner.Controls.Add(this.abteilungszusatz);
            this.Ansprechpartner.Controls.Add(this.departmentAdditionalLbl);
            this.Ansprechpartner.Controls.Add(this.geschäftsbereich);
            this.Ansprechpartner.Controls.Add(this.departmentLabel);
            this.Ansprechpartner.Controls.Add(this.firma);
            this.Ansprechpartner.Controls.Add(this.firmaLbl);
            this.Ansprechpartner.Controls.Add(this.abteilung);
            this.Ansprechpartner.Controls.Add(this.abteilungLbl);
            this.Ansprechpartner.Controls.Add(this.email);
            this.Ansprechpartner.Controls.Add(this.emailLbl);
            this.Ansprechpartner.Controls.Add(this.telefon);
            this.Ansprechpartner.Controls.Add(this.telefonLbl);
            this.Ansprechpartner.Controls.Add(this.nachname);
            this.Ansprechpartner.Controls.Add(this.nachnameLbl);
            this.Ansprechpartner.Controls.Add(this.vorname);
            this.Ansprechpartner.Controls.Add(this.vornameLbl);
            this.Ansprechpartner.Controls.Add(this.frauRadioBtn);
            this.Ansprechpartner.Controls.Add(this.herrRadioBtn);
            this.Ansprechpartner.Location = new System.Drawing.Point(4, 22);
            this.Ansprechpartner.Name = "Ansprechpartner";
            this.Ansprechpartner.Padding = new System.Windows.Forms.Padding(3);
            this.Ansprechpartner.Size = new System.Drawing.Size(947, 497);
            this.Ansprechpartner.TabIndex = 0;
            this.Ansprechpartner.Text = "Ansprechpartner Erstellen";
            this.Ansprechpartner.UseVisualStyleBackColor = true;
            // 
            // createCustomerBtn
            // 
            this.createCustomerBtn.Location = new System.Drawing.Point(330, 406);
            this.createCustomerBtn.Name = "createCustomerBtn";
            this.createCustomerBtn.Size = new System.Drawing.Size(269, 40);
            this.createCustomerBtn.TabIndex = 22;
            this.createCustomerBtn.Text = "Neu Kunden erstellen";
            this.createCustomerBtn.UseVisualStyleBackColor = true;
            this.createCustomerBtn.Click += new System.EventHandler(this.CreateCustomerBtnClicked);
            // 
            // ort
            // 
            this.ort.Location = new System.Drawing.Point(632, 324);
            this.ort.Name = "ort";
            this.ort.Size = new System.Drawing.Size(212, 20);
            this.ort.TabIndex = 21;
            // 
            // ortLbl
            // 
            this.ortLbl.AutoSize = true;
            this.ortLbl.Location = new System.Drawing.Point(528, 327);
            this.ortLbl.Name = "ortLbl";
            this.ortLbl.Size = new System.Drawing.Size(21, 13);
            this.ortLbl.TabIndex = 20;
            this.ortLbl.Text = "Ort";
            // 
            // strasse
            // 
            this.strasse.Location = new System.Drawing.Point(632, 286);
            this.strasse.Name = "strasse";
            this.strasse.Size = new System.Drawing.Size(212, 20);
            this.strasse.TabIndex = 19;
            // 
            // strasseLbl
            // 
            this.strasseLbl.AutoSize = true;
            this.strasseLbl.Location = new System.Drawing.Point(528, 289);
            this.strasseLbl.Name = "strasseLbl";
            this.strasseLbl.Size = new System.Drawing.Size(42, 13);
            this.strasseLbl.TabIndex = 18;
            this.strasseLbl.Text = "Strasse";
            // 
            // abteilungszusatz
            // 
            this.abteilungszusatz.Location = new System.Drawing.Point(632, 249);
            this.abteilungszusatz.Name = "abteilungszusatz";
            this.abteilungszusatz.Size = new System.Drawing.Size(212, 20);
            this.abteilungszusatz.TabIndex = 17;
            // 
            // departmentAdditionalLbl
            // 
            this.departmentAdditionalLbl.AutoSize = true;
            this.departmentAdditionalLbl.Location = new System.Drawing.Point(528, 252);
            this.departmentAdditionalLbl.Name = "departmentAdditionalLbl";
            this.departmentAdditionalLbl.Size = new System.Drawing.Size(86, 13);
            this.departmentAdditionalLbl.TabIndex = 16;
            this.departmentAdditionalLbl.Text = "Abteilungszusatz";
            // 
            // geschäftsbereich
            // 
            this.geschäftsbereich.Location = new System.Drawing.Point(632, 211);
            this.geschäftsbereich.Name = "geschäftsbereich";
            this.geschäftsbereich.Size = new System.Drawing.Size(212, 20);
            this.geschäftsbereich.TabIndex = 15;
            // 
            // departmentLabel
            // 
            this.departmentLabel.AutoSize = true;
            this.departmentLabel.Location = new System.Drawing.Point(528, 214);
            this.departmentLabel.Name = "departmentLabel";
            this.departmentLabel.Size = new System.Drawing.Size(90, 13);
            this.departmentLabel.TabIndex = 14;
            this.departmentLabel.Text = "Geschäftsbereich";
            // 
            // firma
            // 
            this.firma.Location = new System.Drawing.Point(632, 174);
            this.firma.Name = "firma";
            this.firma.Size = new System.Drawing.Size(212, 20);
            this.firma.TabIndex = 13;
            // 
            // firmaLbl
            // 
            this.firmaLbl.AutoSize = true;
            this.firmaLbl.Location = new System.Drawing.Point(528, 177);
            this.firmaLbl.Name = "firmaLbl";
            this.firmaLbl.Size = new System.Drawing.Size(32, 13);
            this.firmaLbl.TabIndex = 12;
            this.firmaLbl.Text = "Firma";
            // 
            // abteilung
            // 
            this.abteilung.Location = new System.Drawing.Point(632, 137);
            this.abteilung.Name = "abteilung";
            this.abteilung.Size = new System.Drawing.Size(212, 20);
            this.abteilung.TabIndex = 11;
            // 
            // abteilungLbl
            // 
            this.abteilungLbl.AutoSize = true;
            this.abteilungLbl.Location = new System.Drawing.Point(528, 140);
            this.abteilungLbl.Name = "abteilungLbl";
            this.abteilungLbl.Size = new System.Drawing.Size(51, 13);
            this.abteilungLbl.TabIndex = 10;
            this.abteilungLbl.Text = "Abteilung";
            // 
            // email
            // 
            this.email.Location = new System.Drawing.Point(171, 249);
            this.email.Name = "email";
            this.email.Size = new System.Drawing.Size(212, 20);
            this.email.TabIndex = 9;
            // 
            // emailLbl
            // 
            this.emailLbl.AutoSize = true;
            this.emailLbl.Location = new System.Drawing.Point(94, 252);
            this.emailLbl.Name = "emailLbl";
            this.emailLbl.Size = new System.Drawing.Size(32, 13);
            this.emailLbl.TabIndex = 8;
            this.emailLbl.Text = "Email";
            // 
            // telefon
            // 
            this.telefon.Location = new System.Drawing.Point(171, 211);
            this.telefon.Name = "telefon";
            this.telefon.Size = new System.Drawing.Size(212, 20);
            this.telefon.TabIndex = 7;
            // 
            // telefonLbl
            // 
            this.telefonLbl.AutoSize = true;
            this.telefonLbl.Location = new System.Drawing.Point(94, 214);
            this.telefonLbl.Name = "telefonLbl";
            this.telefonLbl.Size = new System.Drawing.Size(43, 13);
            this.telefonLbl.TabIndex = 6;
            this.telefonLbl.Text = "Telefon";
            // 
            // nachname
            // 
            this.nachname.Location = new System.Drawing.Point(171, 174);
            this.nachname.Name = "nachname";
            this.nachname.Size = new System.Drawing.Size(212, 20);
            this.nachname.TabIndex = 5;
            // 
            // nachnameLbl
            // 
            this.nachnameLbl.AutoSize = true;
            this.nachnameLbl.Location = new System.Drawing.Point(94, 177);
            this.nachnameLbl.Name = "nachnameLbl";
            this.nachnameLbl.Size = new System.Drawing.Size(59, 13);
            this.nachnameLbl.TabIndex = 4;
            this.nachnameLbl.Text = "Nachname";
            // 
            // vorname
            // 
            this.vorname.Location = new System.Drawing.Point(171, 137);
            this.vorname.Name = "vorname";
            this.vorname.Size = new System.Drawing.Size(212, 20);
            this.vorname.TabIndex = 3;
            // 
            // vornameLbl
            // 
            this.vornameLbl.AutoSize = true;
            this.vornameLbl.Location = new System.Drawing.Point(94, 140);
            this.vornameLbl.Name = "vornameLbl";
            this.vornameLbl.Size = new System.Drawing.Size(49, 13);
            this.vornameLbl.TabIndex = 2;
            this.vornameLbl.Text = "Vorname";
            // 
            // frauRadioBtn
            // 
            this.frauRadioBtn.AutoSize = true;
            this.frauRadioBtn.Location = new System.Drawing.Point(171, 94);
            this.frauRadioBtn.Name = "frauRadioBtn";
            this.frauRadioBtn.Size = new System.Drawing.Size(46, 17);
            this.frauRadioBtn.TabIndex = 1;
            this.frauRadioBtn.TabStop = true;
            this.frauRadioBtn.Text = "Frau";
            this.frauRadioBtn.UseVisualStyleBackColor = true;
            // 
            // herrRadioBtn
            // 
            this.herrRadioBtn.AutoSize = true;
            this.herrRadioBtn.Location = new System.Drawing.Point(97, 94);
            this.herrRadioBtn.Name = "herrRadioBtn";
            this.herrRadioBtn.Size = new System.Drawing.Size(45, 17);
            this.herrRadioBtn.TabIndex = 0;
            this.herrRadioBtn.TabStop = true;
            this.herrRadioBtn.Text = "Herr";
            this.herrRadioBtn.UseVisualStyleBackColor = true;
            // 
            // Projekt
            // 
            this.Projekt.Controls.Add(this.projektTitel);
            this.Projekt.Controls.Add(this.projektTitelLbl);
            this.Projekt.Controls.Add(this.createProjectBtn);
            this.Projekt.Controls.Add(this.ansprechpartnerComboBox);
            this.Projekt.Controls.Add(this.projektBeschreibung);
            this.Projekt.Controls.Add(this.disponent);
            this.Projekt.Controls.Add(this.gesprächsperson);
            this.Projekt.Controls.Add(this.verrechnungssatz);
            this.Projekt.Controls.Add(this.anzahlStunden);
            this.Projekt.Controls.Add(this.endDatumDtPikr);
            this.Projekt.Controls.Add(this.startDatumDtPikr);
            this.Projekt.Controls.Add(this.projektBeschreibungLbl);
            this.Projekt.Controls.Add(this.disponentLbl);
            this.Projekt.Controls.Add(this.gesprächspersonLbl);
            this.Projekt.Controls.Add(this.verrechnungssatzLbl);
            this.Projekt.Controls.Add(this.anzahlStundenLbl);
            this.Projekt.Controls.Add(this.ansprechPartnerLbl);
            this.Projekt.Controls.Add(this.endDatumLbl);
            this.Projekt.Controls.Add(this.startDatumLbl);
            this.Projekt.Controls.Add(this.projektnummer);
            this.Projekt.Controls.Add(this.projektnummerLbl);
            this.Projekt.Location = new System.Drawing.Point(4, 22);
            this.Projekt.Name = "Projekt";
            this.Projekt.Padding = new System.Windows.Forms.Padding(3);
            this.Projekt.Size = new System.Drawing.Size(947, 497);
            this.Projekt.TabIndex = 1;
            this.Projekt.Text = "Projekt Erstellen";
            this.Projekt.UseVisualStyleBackColor = true;
            // 
            // createProjectBtn
            // 
            this.createProjectBtn.Location = new System.Drawing.Point(308, 412);
            this.createProjectBtn.Name = "createProjectBtn";
            this.createProjectBtn.Size = new System.Drawing.Size(322, 40);
            this.createProjectBtn.TabIndex = 23;
            this.createProjectBtn.Text = "Neu Projekt erstellen";
            this.createProjectBtn.UseVisualStyleBackColor = true;
            this.createProjectBtn.Click += new System.EventHandler(this.CreateProjectBtnClicked);
            // 
            // ansprechpartnerComboBox
            // 
            this.ansprechpartnerComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ansprechpartnerComboBox.FormattingEnabled = true;
            this.ansprechpartnerComboBox.Location = new System.Drawing.Point(169, 239);
            this.ansprechpartnerComboBox.Name = "ansprechpartnerComboBox";
            this.ansprechpartnerComboBox.Size = new System.Drawing.Size(200, 21);
            this.ansprechpartnerComboBox.TabIndex = 18;
            // 
            // projektBeschreibung
            // 
            this.projektBeschreibung.Location = new System.Drawing.Point(570, 191);
            this.projektBeschreibung.Multiline = true;
            this.projektBeschreibung.Name = "projektBeschreibung";
            this.projektBeschreibung.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.projektBeschreibung.Size = new System.Drawing.Size(317, 161);
            this.projektBeschreibung.TabIndex = 17;
            // 
            // disponent
            // 
            this.disponent.Location = new System.Drawing.Point(570, 109);
            this.disponent.Name = "disponent";
            this.disponent.Size = new System.Drawing.Size(317, 20);
            this.disponent.TabIndex = 16;
            // 
            // gesprächsperson
            // 
            this.gesprächsperson.Location = new System.Drawing.Point(570, 60);
            this.gesprächsperson.Name = "gesprächsperson";
            this.gesprächsperson.Size = new System.Drawing.Size(317, 20);
            this.gesprächsperson.TabIndex = 15;
            // 
            // verrechnungssatz
            // 
            this.verrechnungssatz.Location = new System.Drawing.Point(169, 332);
            this.verrechnungssatz.Name = "verrechnungssatz";
            this.verrechnungssatz.Size = new System.Drawing.Size(200, 20);
            this.verrechnungssatz.TabIndex = 14;
            // 
            // anzahlStunden
            // 
            this.anzahlStunden.Location = new System.Drawing.Point(169, 286);
            this.anzahlStunden.Name = "anzahlStunden";
            this.anzahlStunden.Size = new System.Drawing.Size(200, 20);
            this.anzahlStunden.TabIndex = 13;
            // 
            // endDatumDtPikr
            // 
            this.endDatumDtPikr.Location = new System.Drawing.Point(169, 155);
            this.endDatumDtPikr.Name = "endDatumDtPikr";
            this.endDatumDtPikr.Size = new System.Drawing.Size(200, 20);
            this.endDatumDtPikr.TabIndex = 11;
            // 
            // startDatumDtPikr
            // 
            this.startDatumDtPikr.Location = new System.Drawing.Point(169, 110);
            this.startDatumDtPikr.Name = "startDatumDtPikr";
            this.startDatumDtPikr.Size = new System.Drawing.Size(200, 20);
            this.startDatumDtPikr.TabIndex = 10;
            // 
            // projektBeschreibungLbl
            // 
            this.projektBeschreibungLbl.AutoSize = true;
            this.projektBeschreibungLbl.Location = new System.Drawing.Point(438, 191);
            this.projektBeschreibungLbl.Name = "projektBeschreibungLbl";
            this.projektBeschreibungLbl.Size = new System.Drawing.Size(108, 13);
            this.projektBeschreibungLbl.TabIndex = 9;
            this.projektBeschreibungLbl.Text = "Projekt Beschreibung";
            // 
            // disponentLbl
            // 
            this.disponentLbl.AutoSize = true;
            this.disponentLbl.Location = new System.Drawing.Point(439, 110);
            this.disponentLbl.Name = "disponentLbl";
            this.disponentLbl.Size = new System.Drawing.Size(55, 13);
            this.disponentLbl.TabIndex = 8;
            this.disponentLbl.Text = "Disponent";
            // 
            // gesprächspersonLbl
            // 
            this.gesprächspersonLbl.AutoSize = true;
            this.gesprächspersonLbl.Location = new System.Drawing.Point(439, 64);
            this.gesprächspersonLbl.Name = "gesprächspersonLbl";
            this.gesprächspersonLbl.Size = new System.Drawing.Size(90, 13);
            this.gesprächspersonLbl.TabIndex = 7;
            this.gesprächspersonLbl.Text = "Gesprächsperson";
            // 
            // verrechnungssatzLbl
            // 
            this.verrechnungssatzLbl.AutoSize = true;
            this.verrechnungssatzLbl.Location = new System.Drawing.Point(62, 335);
            this.verrechnungssatzLbl.Name = "verrechnungssatzLbl";
            this.verrechnungssatzLbl.Size = new System.Drawing.Size(92, 13);
            this.verrechnungssatzLbl.TabIndex = 6;
            this.verrechnungssatzLbl.Text = "Verrechnungssatz";
            // 
            // anzahlStundenLbl
            // 
            this.anzahlStundenLbl.AutoSize = true;
            this.anzahlStundenLbl.Location = new System.Drawing.Point(62, 289);
            this.anzahlStundenLbl.Name = "anzahlStundenLbl";
            this.anzahlStundenLbl.Size = new System.Drawing.Size(82, 13);
            this.anzahlStundenLbl.TabIndex = 5;
            this.anzahlStundenLbl.Text = "Anzahl Stunden";
            // 
            // ansprechPartnerLbl
            // 
            this.ansprechPartnerLbl.AutoSize = true;
            this.ansprechPartnerLbl.Location = new System.Drawing.Point(62, 242);
            this.ansprechPartnerLbl.Name = "ansprechPartnerLbl";
            this.ansprechPartnerLbl.Size = new System.Drawing.Size(85, 13);
            this.ansprechPartnerLbl.TabIndex = 4;
            this.ansprechPartnerLbl.Text = "Ansprechpartner";
            // 
            // endDatumLbl
            // 
            this.endDatumLbl.AutoSize = true;
            this.endDatumLbl.Location = new System.Drawing.Point(62, 155);
            this.endDatumLbl.Name = "endDatumLbl";
            this.endDatumLbl.Size = new System.Drawing.Size(60, 13);
            this.endDatumLbl.TabIndex = 3;
            this.endDatumLbl.Text = "End Datum";
            // 
            // startDatumLbl
            // 
            this.startDatumLbl.AutoSize = true;
            this.startDatumLbl.Location = new System.Drawing.Point(62, 110);
            this.startDatumLbl.Name = "startDatumLbl";
            this.startDatumLbl.Size = new System.Drawing.Size(63, 13);
            this.startDatumLbl.TabIndex = 2;
            this.startDatumLbl.Text = "Start Datum";
            // 
            // projektnummer
            // 
            this.projektnummer.Location = new System.Drawing.Point(169, 61);
            this.projektnummer.Name = "projektnummer";
            this.projektnummer.Size = new System.Drawing.Size(48, 20);
            this.projektnummer.TabIndex = 1;
            // 
            // projektnummerLbl
            // 
            this.projektnummerLbl.AutoSize = true;
            this.projektnummerLbl.Location = new System.Drawing.Point(62, 64);
            this.projektnummerLbl.Name = "projektnummerLbl";
            this.projektnummerLbl.Size = new System.Drawing.Size(77, 13);
            this.projektnummerLbl.TabIndex = 0;
            this.projektnummerLbl.Text = "Projectnummer";
            // 
            // Angebot
            // 
            this.Angebot.Controls.Add(this.button1);
            this.Angebot.Controls.Add(this.label1);
            this.Angebot.Controls.Add(this.projektComboBox);
            this.Angebot.Location = new System.Drawing.Point(4, 22);
            this.Angebot.Name = "Angebot";
            this.Angebot.Size = new System.Drawing.Size(947, 497);
            this.Angebot.TabIndex = 2;
            this.Angebot.Text = "Angebot Generieren";
            this.Angebot.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(309, 238);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(314, 45);
            this.button1.TabIndex = 2;
            this.button1.Text = "Angebot generieren";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(161, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Auswahl ein Projekt";
            // 
            // projektComboBox
            // 
            this.projektComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.projektComboBox.FormattingEnabled = true;
            this.projektComboBox.Location = new System.Drawing.Point(368, 91);
            this.projektComboBox.Name = "projektComboBox";
            this.projektComboBox.Size = new System.Drawing.Size(469, 21);
            this.projektComboBox.TabIndex = 0;
            // 
            // backgrdDBWorker
            // 
            this.backgrdDBWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgrdDBWorker_DoWork);
            this.backgrdDBWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgrdDBWorker_RunWorkerCompleted);
            // 
            // projektTitel
            // 
            this.projektTitel.Location = new System.Drawing.Point(570, 152);
            this.projektTitel.Name = "projektTitel";
            this.projektTitel.Size = new System.Drawing.Size(317, 20);
            this.projektTitel.TabIndex = 25;
            // 
            // projektTitelLbl
            // 
            this.projektTitelLbl.AutoSize = true;
            this.projektTitelLbl.Location = new System.Drawing.Point(439, 153);
            this.projektTitelLbl.Name = "projektTitelLbl";
            this.projektTitelLbl.Size = new System.Drawing.Size(63, 13);
            this.projektTitelLbl.TabIndex = 24;
            this.projektTitelLbl.Text = "Projekt Titel";
            // 
            // ContractDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(979, 547);
            this.Controls.Add(this.ContractGenerator);
            this.Name = "ContractDetails";
            this.Text = "Contract Detail";
            this.Load += new System.EventHandler(this.ContractDetails_Load);
            this.ContractGenerator.ResumeLayout(false);
            this.Ansprechpartner.ResumeLayout(false);
            this.Ansprechpartner.PerformLayout();
            this.Projekt.ResumeLayout(false);
            this.Projekt.PerformLayout();
            this.Angebot.ResumeLayout(false);
            this.Angebot.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl ContractGenerator;
        private System.Windows.Forms.TabPage Ansprechpartner;
        private System.Windows.Forms.Button createCustomerBtn;
        private System.Windows.Forms.TextBox ort;
        private System.Windows.Forms.Label ortLbl;
        private System.Windows.Forms.TextBox strasse;
        private System.Windows.Forms.Label strasseLbl;
        private System.Windows.Forms.TextBox abteilungszusatz;
        private System.Windows.Forms.Label departmentAdditionalLbl;
        private System.Windows.Forms.TextBox geschäftsbereich;
        private System.Windows.Forms.Label departmentLabel;
        private System.Windows.Forms.TextBox firma;
        private System.Windows.Forms.Label firmaLbl;
        private System.Windows.Forms.TextBox abteilung;
        private System.Windows.Forms.Label abteilungLbl;
        private System.Windows.Forms.TextBox email;
        private System.Windows.Forms.Label emailLbl;
        private System.Windows.Forms.TextBox telefon;
        private System.Windows.Forms.Label telefonLbl;
        private System.Windows.Forms.TextBox nachname;
        private System.Windows.Forms.Label nachnameLbl;
        private System.Windows.Forms.TextBox vorname;
        private System.Windows.Forms.Label vornameLbl;
        private System.Windows.Forms.RadioButton frauRadioBtn;
        private System.Windows.Forms.RadioButton herrRadioBtn;
        private System.Windows.Forms.TabPage Projekt;
        private System.Windows.Forms.TabPage Angebot;
        private System.Windows.Forms.Label startDatumLbl;
        private System.Windows.Forms.TextBox projektnummer;
        private System.Windows.Forms.Label projektnummerLbl;
        private System.Windows.Forms.Label verrechnungssatzLbl;
        private System.Windows.Forms.Label anzahlStundenLbl;
        private System.Windows.Forms.Label ansprechPartnerLbl;
        private System.Windows.Forms.Label endDatumLbl;
        private System.Windows.Forms.Label projektBeschreibungLbl;
        private System.Windows.Forms.Label disponentLbl;
        private System.Windows.Forms.Label gesprächspersonLbl;
        private System.Windows.Forms.ComboBox ansprechpartnerComboBox;
        private System.Windows.Forms.TextBox projektBeschreibung;
        private System.Windows.Forms.TextBox disponent;
        private System.Windows.Forms.TextBox gesprächsperson;
        private System.Windows.Forms.TextBox verrechnungssatz;
        private System.Windows.Forms.TextBox anzahlStunden;
        private System.Windows.Forms.DateTimePicker endDatumDtPikr;
        private System.Windows.Forms.DateTimePicker startDatumDtPikr;
        private System.Windows.Forms.Button createProjectBtn;
        private System.ComponentModel.BackgroundWorker backgrdDBWorker;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox projektComboBox;
        private System.Windows.Forms.TextBox projektTitel;
        private System.Windows.Forms.Label projektTitelLbl;
    }
}

