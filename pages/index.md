const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, WidthType, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// Helper: normal paragraph
const p = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, size: 24, font: "Times New Roman", ...opts })],
  spacing: { after: 120, line: 360 },
  alignment: AlignmentType.JUSTIFIED,
});

// Helper: bold heading inline
const ph = (text) => new Paragraph({
  children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman" })],
  spacing: { before: 200, after: 100, line: 360 },
  alignment: AlignmentType.JUSTIFIED,
});

// Helper: indented block quote
const blockQuote = (lines) => lines.map(line => new Paragraph({
  children: [new TextRun({ text: line, size: 22, font: "Times New Roman", italics: false })],
  spacing: { after: 80, line: 340 },
  indent: { left: 720 },
  border: {
    left: { style: BorderStyle.SINGLE, size: 4, color: "888888" }
  },
}));

// Helper: reference entry
const ref = (text) => new Paragraph({
  children: [new TextRun({ text, size: 22, font: "Times New Roman" })],
  spacing: { after: 80, line: 340 },
  indent: { left: 720, hanging: 720 },
});

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1800, right: 1800, bottom: 1800, left: 1800 }
      }
    },
    children: [
      // Title
      new Paragraph({
        children: [new TextRun({
          text: "Bildung und Erziehung in der frühkindlichen Bildung in Deutschland",
          bold: true, size: 28, font: "Times New Roman"
        })],
        spacing: { before: 0, after: 300 },
        alignment: AlignmentType.CENTER,
      }),

      // I. Problemhintergrund
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "I. Problemhintergrund", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      ph("1-1. Zur Problematik der Konzepte Bildung und Erziehung"),

      p("In der frühkindlichen Bildung Deutschlands nimmt die Zahl von Kindern mit unterschiedlichen kulturellen Hintergründen infolge wachsender Migration stetig zu (Anm. 1)1). So sind etwa Fachkräfte mit Migrationshintergrund in Kindertageseinrichtungen tätig, und es existieren finanzielle Fördermaßnahmen für den Spracherwerb sowie Projekte zur Ausbildung von Fachkräften für Zugewanderte2)."),

      p("Das frühkindliche Bildungssystem in Deutschland ist im Rahmen des Sozialgesetzbuchs Achtes Buch (SGB VIII: Kinder- und Jugendhilfe) geregelt. Darin wird Bildung, Erziehung und Betreuung eine zentrale Rolle für die soziale, emotionale, körperliche und geistige Entwicklung von Kindern zugeschrieben – sie nehmen eine bedeutende Stellung in der frühkindlichen Bildung in Deutschland ein. Bildung und Erziehung tragen beide die Bedeutung von ‚Erziehung' bzw. ‚education', doch wurden beide Konzepte bislang ausführlich diskutiert3)4). In englischsprachigen Fachbeiträgen werden Bildung und Erziehung häufig auf Deutsch belassen, während lediglich Betreuung als „care" übersetzt wird5)6). Teils lassen sich Erziehung und Bildung nur schwer in andere Sprachen übertragen7)8). Daher kann man sagen, dass diese Konzepte für ausländische Fachkräfte, die kein Deutsch als Muttersprache sprechen, inhaltlich schwer verständlich sein können."),

      p("Die vorliegende Studie richtet ihren Blick auf Bildung und Erziehung in der frühkindlichen Bildung Deutschlands."),

      ph("1-2. Diskussionen zu den Konzepten Bildung und Erziehung in der frühkindlichen Praxis"),

      p("Auf Japanisch wird Bildung unter anderem mit 教育 (Bildung, Menschenformung)9) und 陶冶 (Kultivierung)10) übersetzt, Erziehung hingegen mit 教育 oder 訓育 (Disziplinierung)11) – beide Begriffe sind vieldeutig12). Was das Verständnis und die Bedeutung der Konzepte Bildung und Erziehung in der jüngeren frühkindlichen Bildung betrifft, so befasst sich Nakanishi (2013, 2014)14)15) in japanischsprachigen Vorarbeiten mit der frühkindlichen Bildung in Deutschland. Sie hebt dabei insbesondere Gerd Schäfers Bildungsverständnis hervor und beschreibt es als Selbstbildung – als einen Prozess der inneren Welterschließung des Kindes. Als Forschungsdesiderat benennt sie die mangelnde Konkretisierung für die Praxis. Zudem werden zwar der nordrhein-westfälische Bildungsplan thematisiert16), nicht jedoch die amtlichen Dokumente aller 16 Bundesländer. Banno (2017)17) stellt die Vielfalt der Bildungspläne der einzelnen Bundesländer in der deutschen frühkindlichen Bildung dar und unterscheidet grob drei Ansätze – Ko-Konstruktion, den situationsorientierten und den Selbstbildungsansatz –, weist jedoch darauf hin, dass die Klassifikation der Bildungspläne je nach Schwerpunktsetzung variiert. Eine gezielte Auseinandersetzung mit den Konzepten Bildung und Erziehung sowie mit der frühkindlichen Praxis fehlt dabei. Zur Erziehung liegt zwar ein älterer Beitrag von Ogasawara (1970)18) vor, der die wissenschaftliche Pädagogik anhand der Zeitschrift Die Erziehung behandelt, doch ist eine systematische Aufarbeitung des Erziehungsbegriffs in der jüngeren japanischsprachigen Forschung zur frühkindlichen Bildung noch unzureichend."),

      p("In Deutschland gibt es den ‚Gemeinsamen Rahmen der Länder für die frühe Bildung in Kindertageseinrichtungen'19), der als gemeinsame Bildungsgrundlage für Kitas und Schulen fungiert (im Folgenden: Rahmen). In diesem Rahmen wird ausdrücklich darauf hingewiesen, dass Bildung und Erziehung in der frühkindlichen Praxis nicht klar voneinander getrennt werden, da sie als untrennbare Wechselwirkung verstanden werden."),

      p("Dem Rahmen20) zufolge sind Bildung und Erziehung in der frühkindlichen Bildung (frühkindliche Bildung) als Zielsetzung untrennbar miteinander verbunden. Im Mittelpunkt der Bildungsbemühungen steht die Stärkung des kindlichen Selbstkonzepts. Erziehung bezeichnet dabei vor allem die Einflussnahme auf das Verhalten und die Sozialisation des Kindes durch Unterstützung, Begleitung und Anregung seitens anderer – in der Regel Erwachsener. Diese unterschiedlichen Bedeutungsauslegungen von Bildung und Erziehung werden zwar sichtbar gemacht, bleiben jedoch auf einer konzeptuellen und abstrakten Ebene, sodass konkrete Praxissituationen nur schwer vorstellbar sind."),

      p("Zur Geschichte der frühkindlichen Bildungspraxis stellen Gauß & Wollnitz (2023)21) fest, dass Deutschland auf sehr unterschiedliche historische Erfahrungen in Ost- und Westdeutschland zurückblickt (Anm. 2) und ein Wandel von familienzentrierter hin zu öffentlicher Bildung und Betreuung stattgefunden hat. Dabei war Betreuung – die körperliche und emotionale Fürsorge für das Kind – lange vorrangig Aufgabe der Familie, Erziehung Aufgabe der Kita und Bildung Aufgabe der Schule. Heute hingegen gilt Bildung als von Beginn an eng mit Erziehung verbunden; in der Frühpädagogik wird die Wechselwirkung von Bildung, Erziehung und Betreuung betont, und diese Konzepte werden nicht mehr voneinander getrennt betrachtet. Der frühkindliche Lernprozess wird als komplex, aktiv und in sozialen Beziehungen eingebettet verstanden, direkt verknüpft mit den Erfahrungen des Kindes. Alle drei Elemente beginnen demnach bei der Geburt in der Familie, werden durch vorschulische Einrichtungen unterstützt und ergänzt und in der Schule fortgesetzt. Die historischen Unterschiede zwischen Ost und West sowie der Wandel im Verständnis von Bildung und Erziehung im Laufe der Zeit werden damit deutlich. Deutschland ist ein Bundesstaat; die Praxis richtet sich nach den jeweiligen Landesgesetzen und Bildungsplänen der 16 Bundesländer. Wie Bildung und Erziehung in den aktuellen Landesgesetzen und Bildungsplänen der einzelnen Länder konkret beschrieben werden, ist bislang nicht hinreichend untersucht worden."),

      p("Frindte & Mierendorff (2017)22) zeigen, dass Bildung und Erziehung in der deutschen frühkindlichen Bildung das theoretische Fundament des spezifisch deutschen Bildungsverständnisses bilden und als Schlüssel zum Verständnis von System und Praxis gelten. Unter Berücksichtigung des historischen Kontexts benennen sie drei konzeptionelle Positionen zu Bildung und Erziehung seit den 2000er-Jahren, als Bildung zum zentralen Begriff der frühkindlichen Bildung avancierte: Erstens Gerd Schäfer, für den Bildung ein von Kindern selbst gestalteter Prozess der Selbstbildung ist. Zweitens Wassilios Fthenakis, der Bildung als Ko-Konstruktion begreift – als gemeinsam im sozialen Austausch entstehendes Lernen. Drittens Laewen, der Bildung und Erziehung unterscheidet: Bildung als eigentätige Weltaneignung des Kindes, Erziehung als Aktivität Erwachsener, um kindliche Kräfte zu entfalten. Zudem weisen sie darauf hin, dass Bildung in der deutschen frühkindlichen Bildung ursprünglich den Prozess meint, durch den sich Kinder aktiv mit der Welt auseinandersetzen und sich selbst formen, und dass dieses Verständnis im Verhältnis zur Erziehung entsteht – in jüngster Zeit jedoch durch Institutionalisierung und ökonomische Anforderungen eingeengt zu werden droht. Die Konzepte unterliegen damit einem zeitbedingten Wandel und sind von erheblicher Komplexität."),

      p("Pestalozzi-Fröbel-Verband e.V. (2023)23) hat untersucht, wie theoretische und empirische Debatten in die amtlichen Dokumente der Bildungspläne (Bildungsplan u. a.; im Folgenden: Bildungspläne) einfließen, in welchem Maß sie in die Landesregelungen integriert werden und ob wissenschaftliche Erkenntnisse Berücksichtigung finden. Die Bildungspläne aller 16 Bundesländer wurden analysiert und ihre jeweiligen Besonderheiten herausgearbeitet. Allerdings wurden die Bildungspläne von 9 der 16 Länder nach 2024 überarbeitet, sodass eine Aufarbeitung auf Grundlage der aktuellen Bildungspläne erforderlich erscheint. Zudem existieren in jedem Bundesland landeseigene Gesetze (Landesgesetz; KiTa-Gesetz u. a.; im Folgenden: Landesgesetze), von denen 12 nach 2024 überarbeitet wurden – auch diese Landesgesetze sollten daher überblicksartig betrachtet werden."),

      p("Die Konzepte Bildung und Erziehung sind somit vielschichtig und werden vor dem Hintergrund ihrer historischen Entwicklung diskutiert. Für die tatsächliche Praxis in deutschen frühkindlichen Einrichtungen sind jedoch die aktuellen Landesgesetze und Bildungspläne der jeweiligen Bundesländer maßgeblich (Anm. 1)24)."),

      p("Es besteht daher die Notwendigkeit, die in den vergangenen Jahren überarbeiteten Landesgesetze und Bildungspläne im Hinblick auf Bildung und Erziehung erneut zu sichten. Zudem gilt es, die Konzepte von Bildung und Erziehung, ihre jeweilige Bedeutung und Verortung sowie ihre Beziehung zueinander im konkreten Kontext der frühkindlichen Bildungspraxis verständlich und nachvollziehbar darzustellen – auch für nicht-muttersprachliche Fachkräfte."),

      ph("1-3. Ziel der vorliegenden Studie"),

      p("Vor diesem Hintergrund verfolgt die vorliegende Studie das Ziel, auf der Grundlage der amtlichen Dokumente – der Landesgesetze sowie der Bildungspläne aller 16 Bundesländer – zu klären, als welche konkreten Konzepte Bildung und Erziehung in der frühkindlichen Bildung in Deutschland verankert sind. Darüber hinaus sollen die Beziehung und die Merkmale beider Konzepte aus der Perspektive der Praxis herausgearbeitet werden. Damit soll ein Beitrag geleistet werden, um Bildung und Erziehung im Kontext konkreter frühpädagogischer Praxissituationen greifbar zu machen."),

      // II. Methode
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "II. Methode", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      p("Aus den amtlichen Dokumenten der Landesgesetze sowie der Bildungspläne aller 16 Bundesländer werden Passagen extrahiert und systematisch aufbereitet, die das Verständnis und die Bedeutung der Konzepte Bildung und Erziehung betreffen. Das methodische Vorgehen gliedert sich wie folgt: Zunächst werden in den jeweiligen Landesgesetzen und Bildungsplänen die Begriffe Bildung und Erziehung recherchiert und herausgearbeitet (Abb. 1). Anschließend werden Passagen, die konkrete Aussagen oder Interpretationen zu den Konzepten sowie formulierte Ziele enthalten – also Antworten auf Fragen wie „Was ist Bildung?" und „Was ist Erziehung?" –, exzerpiert und übersetzt. Abschließend erfolgt eine übergreifende Analyse aller 16 Bundesländer."),

      p("In der vorliegenden Studie werden nur die Passagen, die konkrete Angaben zu Bildung und Erziehung enthalten, in einem eingerahmten Kasten dargestellt. In Klammern wird dabei angegeben, ob es sich um ein Zitat aus einem Landesgesetz oder einem Bildungsplan handelt und welchem Bundesland die Passage entstammt. An Stellen, an denen die japanische Übersetzung allein den Sinn nicht hinreichend wiedergibt, werden ergänzende deutsche Begriffe eingefügt."),

      // III. Ergebnisse und Diskussion
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "III. Ergebnisse und Diskussion", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      p("In der vorliegenden Studie wurden aus allen 16 Landesgesetzen und Bildungsplänen ausschließlich jene Passagen herausgegriffen, die Aussagen zu den Konzepten Bildung und Erziehung enthalten. Dabei zeigte sich, dass unter den Landesgesetzen lediglich ‹1› Nordrhein-Westfalen konkrete Angaben enthält. Bei den Bildungsplänen wiesen 【1】 Baden-Württemberg, 【2】 Mecklenburg-Vorpommern, 【3】 Schleswig-Holstein und 【4】 Thüringen spezifische Ausführungen zu den Konzepten Bildung und Erziehung auf."),

      p("Die Auswertung dieser Dokumente erbrachte zwei zentrale Befunde: Erstens wurden Bildung und Erziehung – übereinstimmend mit dem SGB VIII und dem Rahmen – nicht als voneinander getrennte, sondern als eigenständige und zugleich untrennbar aufeinander bezogene Konzepte verstanden. Zudem werden sie in allen Bundesländern einheitlich im Verbund „Bildung, Erziehung und Betreuung" behandelt. Zweitens wurde deutlich: Bildung meint die Selbstbildung, durch die Kinder ihre Welt und sich selbst aktiv gestalten, und entsteht im alltäglichen Erleben in der frühkindlichen Bildung. Erziehung hingegen geht von Erwachsenen aus, begleitet die Bildungsprozesse und ist im Kontext der frühkindlichen Bildung verankert. Diese Befunde werden im Folgenden anhand konkreter Textbelege erläutert und diskutiert."),

      ph("3-1. Eigenständigkeit und wechselseitige Verschränkung von Bildung und Erziehung"),

      p("In allen 16 Landesgesetzen und Bildungsplänen wird „Bildung, Erziehung und Betreuung" als untrennbarer Verbund behandelt, und die Einrichtungen werden als Orte beschrieben, die genau dies leisten – eine Gemeinsamkeit, die sich auch im SGB VIII und im Rahmen findet. Zudem wird festgehalten, dass „Bildung, Erziehung und Betreuung" die Familie ergänzen und Eltern darauf einen Rechtsanspruch haben (vgl. z. B. Berlin, Bayern)."),

      p("【1】 Baden-Württemberg (BW) beschreibt die besondere Eigenständigkeit von Bildung und Erziehung in Deutschland:"),

      ...blockQuote([
        "Kapitel 1  Grundlagen und Zielorientierungen (Bildungsplan · 【1】 BW)",
        "",
        "1.1.2 Bildung, Erziehung und Betreuung",
        "",
        "In Deutschland besitzen die Konzepte ‚Bildung' und ‚Erziehung' eine eigenständige Tradition. Dies liegt auch darin begründet, dass die im deutschsprachigen Raum übliche Unterscheidung zwischen ‚Bildung' und ‚Erziehung' in anderen Sprachen kaum anzutreffen ist.",
        "",
        "In jüngerer Zeit hat sich – besonders im Bereich der Kindertagesbetreuung – ein Rahmen herausgebildet, der ‚Bildung, Erziehung und Betreuung' (die sog. Trias) als zusammengehörige Einheit beschreibt (vgl. § 22 Abs. 3 SGB VIII).",
      ]),

      p("Im Bildungsplan von 【2】 Mecklenburg-Vorpommern (MV), der ebenfalls konkrete Ausführungen zu Bildung und Erziehung enthält, wird als Grundaussage festgehalten, dass Bildung und Erziehung integriert und nicht voneinander zu trennen sind:"),

      ...blockQuote([
        "Kapitel 1  Bildung, Erziehung und Betreuung in Kindertageseinrichtungen (Bildungsplan · 【2】 MV)",
        "",
        "Erziehung und Bildung",
        "",
        "Erziehung und Bildung sind nicht voneinander zu trennen. Sie sind „zwei Seiten derselben Medaille" und bedingen sich gegenseitig. Bildung bezeichnet die Perspektive der Kinder – d. h., was Kinder selbst tun, um Fähigkeiten zu erwerben. Erziehung hingegen bezeichnet die Perspektive der Erwachsenen – d. h., was pädagogische Fachkräfte tun, um kindliche Bildungsprozesse zu ermöglichen und zu unterstützen.",
      ]),

      p("Auf diese Weise wurden Bildung und Erziehung als untrennbar aufeinander bezogene Konzepte verstanden. Dabei ist kennzeichnend, dass Bildung auf die Perspektive des Kindes und Erziehung auf die Perspektive der Erwachsenen bezogen wird. Im Kontext der frühkindlichen Bildung lässt sich zudem schlussfolgern, dass Bildung und Erziehung als ergänzende Leistungen der Familie integriert verwirklicht werden."),

      p("[Tabelle 1: Übersicht der Landesgesetze und Bildungspläne]"),

      ph("3-2. Im Alltag entstehende Bildung und die sie begleitende Erziehung (Bildungsprozesse)"),

      p("Von den 16 Landesgesetzen enthielt lediglich ‹1› Nordrhein-Westfalen (NRW) konkrete Erläuterungen zu Bildung und Erziehung. Bildung wird dort als Selbstbildung verstanden – als aktive Auseinandersetzung mit der eigenen Umgebung. Im Folgenden wird der relevante Abschnitt exzerpiert und diskutiert:"),

      ...blockQuote([
        "§ 15 Frühkindliche Bildung (Landesgesetz · ‹1› NRW)",
        "",
        "Bildung ist die aktive Auseinandersetzung des Kindes mit seiner Umgebung auf der Grundlage seiner bisherigen Lebenserfahrungen. Sie ist ein konstruktiver Prozess, in dem Selbstbildung einerseits durch unmittelbare Wahrnehmung und aktives, experimentierendes Handeln entsteht, andererseits durch den Einfluss der Umgebung, wobei beide Seiten in Wechselwirkung stehen. Bildung wirkt darauf hin, die Entwicklung des Kindes zu einer eigenständigen Persönlichkeit sowie den Erwerb sozialer Kompetenz zu fördern.",
      ]),

      p("Damit wird das Konzept Bildung in einem Teil seiner Inhalte konkret beschrieben. Die alleinige Darstellung des Bildungsbegriffs zeigt, dass dieser einer Erklärung bedarf. Zugleich wird deutlich, dass der Entwicklung einer eigenständigen Persönlichkeit (eigenständige Persönlichkeit) und dem Erwerb sozialer Kompetenz durch Bildung ein hoher Stellenwert beigemessen wird."),

      p("Zum Bildungsbegriff gibt der Bildungsplan von 【2】 Schleswig-Holstein (SH) ebenfalls konkrete Ausführungen. Selbstbildung wird darin als aktive Tätigkeit des Subjekts verstanden; Bildung bedeutet stets ‚sich bilden' – und gerade nicht eine von außen herangetragene Einwirkung. Dass alle Kinder bereits mit der Geburt beginnen, die Welt zu entdecken, ist ein charakteristisches Merkmal von Bildung. Zudem wird betont, dass Erziehung neben Bildung unabdingbar ist, und die frühkindliche Bildung zielt wesentlich darauf ab, die Bildungsprozesse des Kindes zu unterstützen:"),

      ...blockQuote([
        "Kapitel 1  Grundlagen – Bildung, Erziehung und Betreuung in Kindertageseinrichtungen und der Kindertagespflege (Bildungsplan · 【3】 SH)",
        "",
        "1.1.1 Bildung als Aneignungstätigkeit",
        "",
        "In der Kindheitspädagogik wird Bildung als aktive Tätigkeit des Subjekts verstanden und bedeutet stets ‚sich bilden'. Bildung wird nicht durch Erziehung von außen hergestellt, kann jedoch durch Erziehung unterstützt werden – und muss es. Alle Kinder beginnen mit der Geburt, die Welt zu entdecken, und setzen sich mit den vielfältigen Themen ihrer Umwelt kompetent, aktiv und neugierig auseinander. Deshalb wird Bildung auch als Selbstbildung bezeichnet. (…)",
        "",
        "1.1.2 Erziehung als ziel- und werteorientiertes Handeln der Fachkräfte",
        "",
        "Damit Kinder sich bilden können, ist Erziehung notwendig. Diese Notwendigkeit gründet in der anthropologischen Annahme, dass der Mensch von Natur aus erziehungsbedürftig ist. Das bedeutet: Alle Kinder müssen von Erwachsenen in ihren Grundbedürfnissen befriedigt, mit Orientierung versehen sowie in ihren eigenen Bildungsprozessen unterstützt und begleitet werden.",
      ]),

      p("Anschließend wird das Bildungskonzept von 【1】 Mecklenburg-Vorpommern (MV) erläutert. Ähnlich wie in Schleswig-Holstein wird Bildung als Verhältnis des Subjekts zur Welt und als Selbstbildung verstanden. Zudem wird der Unterschied zwischen Bildung und Erziehung ausgeführt: Während Bildungsprozesse lebenslang andauern, richtet sich Erziehung auf die Mündigkeit des Menschen. In 【3】 Thüringen (TH) wird Bildung und Erziehung ebenfalls unterschieden: Bildung bedeutet die Verbindung unserer selbst mit der Welt, Erziehung stimuliert diese Bildungsprozesse. Aus Mecklenburg-Vorpommern (MV) stammen zudem konkrete Beschreibungen der Orte, an denen Bildung stattfindet:"),

      ...blockQuote([
        "Kapitel 1  Bildung, Erziehung und Betreuung in Kindertageseinrichtungen (Bildungsplan · 【1】 MV)",
        "",
        "1. Bildung",
        "",
        "Im Kontext der Früh-, Elementar- bzw. Kindheitspädagogik wird Bildung zunächst allgemein als Verhältnis des Menschen (Subjekt) zur Welt (Verhältnis zur Welt) verstanden. (…) Der Begriff Bildung meint stets Selbstbildung. Bildung bezieht sich auf den einzelnen Menschen und umfasst alle Dimensionen des Lebens – Essen, Schlafen, Spielen, Arbeiten – und entfaltet sich in der Beziehung zur materialen und sozialen Umwelt. (…) Bildung vollzieht sich zwar im Prozess der Weltaneignung und des Lernens, bezeichnet aber weniger die Aneignung bestimmter Inhalte als die ‚Fähigkeit, sich lernend die Welt zu erschließen'. Wesentlich ist dabei: offen sein für neue Welterfahrungen, die Welt als Herausforderung begreifen und selbsttätig, doch im Dialog mit anderen, Lern- und Bildungsprozesse gestalten. (…) Sie beginnt in der Familie und wird in Kindertageseinrichtungen und der Kindertagespflege fortgeführt, ist jedoch nicht auf formale und non-formale Lernorganisationen beschränkt, sondern wurzelt auch in informellen Settings.",
        "",
        "2. Erziehung",
        "",
        "Erziehung ist neben Bildung ein ‚Kernbegriff der Pädagogik' und kann wie folgt verstanden werden. (…) Während in der Schule der hier beschriebene Prozess als ‚Unterricht' bezeichnet wird, heißt er in Kindertageseinrichtungen und der Kindertagespflege ‚Angebote' oder ‚Projekte'. (…) Ein Unterschied zwischen Bildung und Erziehung betrifft die zeitliche Dimension: Bildungsprozesse dauern ein Leben lang an, während Erziehung ‚letztlich darauf zielt, sich selbst überflüssig zu machen und den Menschen zur Mündigkeit zu führen'.",
      ]),

      ...blockQuote([
        "Kapitel 1  Bildungswissenschaftliche Grundlagen (Bildungsplan · 【3】 TH)",
        "",
        "· Bildung: Bildung ist ‚die Verbindung unseres Ichs mit der Welt' (Humboldt). Damit verbunden ist, dass Kinder und Jugendliche die Welt und ihre eigene Position darin verstehen, darüber reflektieren und verantwortungsvoll handeln können.",
        "",
        "· Erziehung: Durch pädagogisches Handeln werden Bildungsprozesse angeregt.",
        "",
        "· Bildungswelten: Bildung ist an vielfältigen Orten und in vielfältigen Situationen möglich.",
      ]),

      p("Bildung vollzieht sich also in der alltäglichen Lebenswelt der Kinder – beim Essen, Schlafen, Spielen. Erziehung hingegen entsteht, wenn Fachkräfte Aktivitäten (Angebote) oder Projekte planen und gestalten, um Prozesse anzuregen und weiterzuentwickeln."),

      p("Bei der Überarbeitung des Bildungsplans wird auch die begriffliche Klärung von Bildung und Erziehung eigens aufgegriffen. Mecklenburg-Vorpommern hat bei der Revision den Umfang des Dokuments erweitert, um Bildung und Erziehung transparenter zu machen. Dadurch war es möglich, Bildung und Erziehung in einem praxisnäheren Verständnis herauszuarbeiten."),

      // IV. Fazit und Ausblick
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "IV. Fazit und Ausblick", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      p("Die vorliegende Studie hat gezeigt, dass unter den Landesgesetzen Nordrhein-Westfalen und unter den Bildungsplänen Mecklenburg-Vorpommern, Schleswig-Holstein und Thüringen konkrete Ausführungen zu den Konzepten Bildung und Erziehung enthalten. Auf der Grundlage dieser Befunde soll nun ein übergreifendes Verständnis von Bildung und Erziehung in der deutschen frühkindlichen Bildung zusammenfassend dargestellt werden."),

      p("Bildung meint die Selbstbildung, durch die Kinder ihre Welt und sich selbst aktiv formen; die dabei ablaufenden Bildungsprozesse dauern ein Leben lang an. Bildung beginnt mit der Geburt und entsteht im frühpädagogischen Kontext im Alltag – beim Schlafen, Essen oder Spielen. Erziehung hingegen geht von Erwachsenen bzw. Fachkräften aus und zielt auf Mündigkeit. Konkret äußert sie sich etwa in Angeboten oder Projekten, die von Fachkräften geplant und gestaltet werden. Bildung und Erziehung sind dabei untrennbar aufeinander bezogen und bilden eine eigenständige, in sich zusammenhängende Einheit. Darüber hinaus übernehmen beide eine ergänzende Funktion gegenüber der Familie – zur Förderung von Selbstständigkeit und sozialer Kompetenz – und sind als zentrale Konzepte der Unterstützung von Kindern in der frühkindlichen Bildung zu verstehen."),

      p("Im Folgenden werden drei Grenzen der vorliegenden Studie benannt:"),

      p("Erstens: Obwohl ein umfassendes Verständnis von Bildung und Erziehung herausgearbeitet werden konnte, konzentriert sich die Studie vorrangig auf jene Bundesländer, in denen konkrete Beschreibungen dieser Konzepte vorliegen. Wie Pestalozzi-Fröbel-Verband e.V. (2023)23) anmerkt, bestehen erhebliche Unterschiede zwischen den einzelnen Ländern, die nicht nur auf konzeptuelle Verständnisfragen, sondern auch darauf zurückzuführen sind, dass Bildung und Erziehung in den jeweiligen Ausführungsgesetzen und Bildungsplänen nicht klar ausgeführt werden. Die Frage, inwiefern die historischen Hintergründe der einzelnen Bundesländer mit diesen Konzepten zusammenhängen, konnte nicht untersucht werden."),

      p("Zweitens: Es wurde zwar deutlich, dass Bildung, Erziehung und Betreuung als untrennbare Einheit verstanden werden und Bildung und Erziehung einander bedingen – das konkrete Verhältnis zur Betreuung konnte jedoch nicht näher beleuchtet werden. Auch die Beziehung zwischen den angestrebten Zielen Selbstständigkeit und sozialer Kompetenz auf der einen und Bildung und Erziehung auf der anderen Seite wurde nicht systematisch untersucht."),

      p("Drittens: Wie Fachkräfte in der konkreten Praxis Bildung und Erziehung als untrennbare Einheit wahrnehmen und umsetzen, konnte nicht ausreichend erörtert werden. Die Analyse amtlicher Dokumente allein reicht nicht aus, um die Sichtweisen der in den einzelnen Bundesländern tätigen Fachkräfte zu erfassen."),

      p("Als Ertrag der vorliegenden Studie ist festzuhalten, dass auf der Grundlage amtlicher Dokumente ein übergreifendes Verständnis von Bildung und Erziehung erarbeitet und diese Konzepte im Kontext der frühkindlichen Bildung greifbarer gemacht werden konnten. Gleichzeitig bleibt die Frage offen, worin genau die eigenständigen Konzepte Bildung und Erziehung voneinander abweichen. Künftige Forschungen sollten daher die theoretischen Besonderheiten und die Eigenständigkeit von Bildung und Erziehung unter Berücksichtigung der genannten Grenzen noch differenzierter und vielschichtiger in den Blick nehmen."),

      // Anmerkungen
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Anmerkungen", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      p("1) Laut Deutschem Jugendinstitut (DJI, 2020) sprechen ein Fünftel aller Kinder zu Hause eine andere Sprache als Deutsch; in Bundesländern wie Hessen und Berlin ist es sogar ein Drittel."),

      p("2) Gauß & Wollnitz (2023) beschreiben die unterschiedlichen historischen Hintergründe in Ost- und Westdeutschland. In Deutschland galt die Betreuung von Kleinkindern lange als „Privatsache", die hauptsächlich von Müttern oder nahen Verwandten übernommen wurde. In der DDR begann in den 1960er-Jahren der Ausbau der Kinderbetreuung; 1970 waren bereits 65 % der Kinder von 3 Jahren bis zum Schulalter und 24 % der unter Dreijährigen in Betreuung. In Westdeutschland wurde der Kindergarten erst nach 1970 offiziell als Elementarstufe des Bildungssystems anerkannt; davor galt er eher als soziale Nothilfe. Ende der 1980er-Jahre wurden in Westdeutschland 99 % der unter Dreijährigen und 88 % der 3- bis 6-Jährigen von nicht berufstätigen Müttern zu Hause betreut; selbst bei berufstätigen Müttern waren es noch 75 %. Erster Versuch, ein Rechtsanspruch auf einen Kindergartenplatz zu verankern, scheiterte 1989 am Widerstand der Länder; seit dem 1. Januar 1996 haben Kinder ab drei Jahren jedoch einen gesetzlichen Anspruch darauf. Der steigende Anteil hochqualifizierter junger Frauen, die Anforderungen des Arbeitsmarkts nach flexiblen Arbeitskräften sowie die damit verbundene Aufgabe der „Vereinbarkeit von Familie und Beruf" haben die deutsche Politik unter Druck gesetzt, die Kinderbetreuung neu zu gestalten. Als Ursprung des Bildungsbegriffs wird Comenius (1592–1670) angeführt, der den Prozess der Selbstbildung (Bildungsprozess) erörtert und betont hat, dass man nur das formen kann, was sich selbst formt. Das BMFSFJ (2006) hält fest, dass das Recht auf Bildung von Geburt an gilt."),

      // Literatur
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "Literaturverzeichnis", bold: true, size: 28, font: "Times New Roman" })],
        spacing: { before: 300, after: 150 },
      }),

      ref("Olszenka, N. & Riedel, B. (2020). Kulturelle Vielfalt in Kitas: Früh gefördert oder abgehängt? Deutsches Jugendinstitut (DJI). https://www.dji.de/themen/kinderbetreuung/kulturelle-vielfalt-in-kitas.html (Zugriff: 25.04.2025)"),
      ref("Netzwerk IQ. (2025). Employment in early childhood education and care: Opportunities and challenges for migrants. https://netzwerk-iq.de/ (Zugriff: 25.04.2025)"),
      ref("Uljens, M. (2001). On general education as a discipline. Studies in Philosophy and Education, 20(4), 291–301."),
      ref("Biesta, G. (2002). How general can Bildung be? Reflections on the future of a modern educational ideal. Journal of Philosophy of Education, 36(3), 377–390."),
      ref("Uljens, M. (2002). The idea of a universal theory of education: An impossible but necessary project? Journal of Philosophy of Education, 36(3), 353–375."),
      ref("Andresen, S. & Koch, K. (2017). Bildung, Erziehung [education] and care in German early childhood settings: Spotlights on current discourses. Journal of Pedagogy, 8(1), 99–117."),
      ref("Benner, D. (2015). Erziehung und Bildung! Zur Konzeptualisierung eines erziehenden Unterrichts, der bildet. Zeitschrift für Pädagogik, 61(4), 481–496."),
      ref("Taylor, C. A. (2017). Is a posthumanist 'Bildung' possible? Reclaiming the promise of 'Bildung' for contemporary higher education. Higher Education, 74(3), 419–435."),
      ref("Toyoda, K. (2017). Bundesrepublik Deutschland – Entwicklungen in der Kita- und Vorschulbildungsreform nach der Wiedervereinigung. In: Izumi C. (Hrsg.), Warum lernen wir von der weltweiten frühkindlichen Bildung? Minerva Shobo, S. 127–157."),
      ref("Torikou, M. (2011). Was der Begriff der Bildung eröffnet – vor dem Hintergrund der deutschen Kitapolitik. Weißume Kindheitswissenschaftliche Schriftenreihe 4. Shiraume Gakuen Universität Verlag."),
      ref("Nakanishi, S. (2023). Bildung in der frühkindlichen Erziehung Deutschlands – Lernen aus der Sicht des Kindes neu denken. Shumpusha."),
      ref("Ministerium für Kultus, Jugend und Sport Baden-Württemberg. (2025). Orientierungsplan für Bildung und Erziehung in baden-württembergischen Kindertageseinrichtungen und Kindertagespflege."),
      ref("Nakanishi, S. (2013). Untersuchung zum ‚Lernen' des Kindes in der Betreuung – mit Fokus auf Schäfers Bildungsverständnis als Selbstbildungstheorie. Forschungsberichte zur frühkindlichen Bildung, 51(2), 154–162."),
      ref("Nakanishi, S. (2014). Diskussionen über die Konzeptualisierung pädagogischer Aufgaben in der Kinderbetreuung in Deutschland. Erziehungswissenschaftliche Forschung, 81(4), 473–483."),
      ref("Nakanishi, S. (2016). Diskussionen über die Qualität frühkindlicher Lernprozesse in Deutschland. Forschungsberichte zur frühkindlichen Bildung, 54(2), 28–36."),
      ref("Banno, S. (2016). Vorschulische Bildung in Deutschland: Stand und Herausforderungen. Abhandlungen der Pädagogischen Fakultät der Tamagawa Universität, 16, 19–47."),
      ref("Ogasawara, M. (1970). Studien zur wissenschaftlichen Pädagogik in den 1920er Jahren – mit Schwerpunkt auf der Zeitschrift ‚Die Erziehung'. Erziehungswissenschaftliche Forschung, 37(1), 1–13."),
      ref("Jugendministerkonferenz/Kultusministerkonferenz (JMK/KMK). (2022). Gemeinsamer Rahmen der Länder für die frühe Bildung in Kindertageseinrichtungen. https://www.kmk.org/ (Zugriff: 20.04.2025)"),
      ref("Gauß, S. & Wollnitz, T. (2023). Die Implementierung des Bildungsbegriffs in der Frühpädagogik. KiTa Fachtexte."),
      ref("Frindte, A. & Mierendorff, J. (2017). Bildung, Erziehung [education] and care in German early childhood settings: Spotlights on current discourses. Journal of Pedagogy, 8(1), 99–120."),
      ref("Pestalozzi-Fröbel-Verband e.V. (2023). Rethinking frühkindliche „Erziehung, Bildung und Betreuung": Fachwissenschaftliche und rechtliche Vermessungen zum Bildungsanspruch in der Kindertagesbetreuung."),

      // Landesgesetze
      new Paragraph({
        children: [new TextRun({ text: "Landesgesetze", bold: true, size: 24, font: "Times New Roman" })],
        spacing: { before: 200, after: 80 },
      }),

      ref("Land Baden-Württemberg. (2009). Gesetz über die Betreuung und Förderung von Kindern in Kindergärten, anderen Tageseinrichtungen und der Kindertagespflege (KiTaG)."),
      ref("Freistaat Bayern. (2005). Bayerisches Kinderbildungs- und -betreuungsgesetz (BayKiBiG)."),
      ref("Land Berlin. (2005). Kindertagesförderungsgesetz (KitaFöG)."),
      ref("Land Brandenburg. (2004). Gesetz über die Entwicklung, Förderung und Betreuung von Kindern in Tageseinrichtungen und in Tagespflege (KitaG)."),
      ref("Freie Hansestadt Bremen. (2000). Bremisches Gesetz zur Förderung von Kindern in Tageseinrichtungen und in Tagespflege (BremKTG)."),
      ref("Freie und Hansestadt Hamburg. (2002). Hamburgisches Kinderbetreuungsgesetz (KiBeG)."),
      ref("Land Hessen. (2005). Hessisches Kinder- und Jugendhilfegesetzbuch (HKJGB)."),
      ref("Land Mecklenburg-Vorpommern. (2004). Gesetz zur Förderung von Kindern in Kindertageseinrichtungen und in Kindertagespflege (KiföG M-V)."),
      ref("Land Niedersachsen. (2005). Gesetz über Tageseinrichtungen für Kinder (KiTaG)."),
      ref("Land Nordrhein-Westfalen. (2008). Kinderbildungsgesetz (KiBiz)."),
      ref("Land Rheinland-Pfalz. (2019). Landesgesetz über die Erziehung, Bildung und Betreuung von Kindern in Tageseinrichtungen und in Kindertagespflege (KiTaG)."),
      ref("Land Saarland. (2004). Gesetz zur Förderung von Kindern in Kindertageseinrichtungen (SBKitaG)."),
      ref("Freistaat Sachsen. (2005). Gesetz über Kindertageseinrichtungen (SächsKitaG)."),
      ref("Land Sachsen-Anhalt. (2013). Gesetz zur Förderung und Betreuung von Kindern in Tageseinrichtungen und in Tagespflege (KiFöG)."),
      ref("Land Schleswig-Holstein. (2020). Kindertagesstättengesetz (KiTaG)."),
      ref("Freistaat Thüringen. (2005). Thüringer Gesetz über die Bildung, Erziehung und Betreuung von Kindern in Tageseinrichtungen und in Tagespflege (ThürKitaG)."),

      // Bildungspläne
      new Paragraph({
        children: [new TextRun({ text: "Bildungspläne", bold: true, size: 24, font: "Times New Roman" })],
        spacing: { before: 200, after: 80 },
      }),

      ref("Ministerium für Kultus, Jugend und Sport Baden-Württemberg. (2025). Orientierungsplan für Bildung und Erziehung in baden-württembergischen Kindertageseinrichtungen und Kindertagespflege."),
      ref("Staatsinstitut für Frühpädagogik und Medienkompetenz. (2024). Der Bayerische Bildungs- und Erziehungsplan für Kinder in Tageseinrichtungen bis zur Einschulung."),
      ref("Senatsverwaltung für Bildung, Jugend und Familie Berlin. (2014). Berliner Bildungsprogramm für Kitas und Kindertagespflege."),
      ref("Ministerium für Bildung, Jugend und Sport des Landes Brandenburg. (2024). Bildungsplan für Kindertagesstätten im Land Brandenburg."),
      ref("Die Senatorin für Kinder und Bildung Bremen. (2017). Rahmenplan für Bildung und Erziehung im Elementarbereich."),
      ref("Freie und Hansestadt Hamburg. (2024). Bildungsleitlinien für die frühkindliche Bildung in Kindertageseinrichtungen."),
      ref("Hessisches Ministerium für Soziales und Integration. (2025). Bildungs- und Erziehungsplan für Kinder von 0 bis 10 Jahren in Hessen."),
      ref("Ministerium für Bildung und Kindertagesförderung Mecklenburg-Vorpommern. (2025). Bildungskonzeption für Kinder von 0 bis 10 Jahren."),
      ref("Niedersächsisches Kultusministerium. (2023). Orientierungsplan für Bildung und Erziehung im Elementarbereich."),
      ref("Ministerium für Kinder, Familie, Flüchtlinge und Integration des Landes Nordrhein-Westfalen. (2018). Bildungsgrundsätze für Kinder von 0 bis 10 Jahren in Kindertagesbetreuung und Schulen im Primarbereich."),
      ref("Ministerium für Bildung Rheinland-Pfalz. (2024). Bildungs- und Erziehungsempfehlungen für Kindertagesstätten in Rheinland-Pfalz."),
      ref("Ministerium für Bildung und Kultur des Saarlandes. (2018). Bildungsprogramm für saarländische Kindertageseinrichtungen."),
      ref("Sächsisches Staatsministerium für Kultus. (2011). Der sächsische Bildungsplan."),
      ref("Ministerium für Arbeit, Soziales, Gesundheit und Gleichstellung des Landes Sachsen-Anhalt. (2025). Bildungsprogramm für Kindertageseinrichtungen in Sachsen-Anhalt."),
      ref("Ministerium für Soziales, Jugend, Familie, Senioren, Integration und Gleichstellung Schleswig-Holstein. (2026). Erfolgreich starten – Leitlinien für Bildung und Erziehung in Kindertageseinrichtungen und der Kindertagespflege Schleswig-Holstein."),
      ref("Thüringer Ministerium für Bildung, Jugend und Sport. (2019). Der Thüringer Bildungsplan bis 18 Jahre."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Bildung_Erziehung_Fruehpaedagogik_DE.docx', buffer);
  console.log('Done!');
}).catch(e => { console.error(e); process.exit(1); });
