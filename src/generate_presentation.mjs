import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import PptxGenJS from "pptxgenjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");
const ASSET_DIR = path.join(ROOT, "assets", "generated");
const OUTPUT_DIR = path.join(ROOT, "output");

const palette = {
  ink: "3A3D46",
  muted: "6B7280",
  line: "E8EDF2",
  white: "FFFFFF",
  blue: "CFEFFF",
  pink: "FFD9E8",
  yellow: "FFF7C7",
  slate: "EEF3F7",
  shadow: "D8DEE8",
};

const slidesMeta = [];

const assets = {
  "title-school-illustration.svg": `<?xml version="1.0" encoding="UTF-8"?>
<svg width="1600" height="980" viewBox="0 0 1600 980" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="1600" height="980" rx="42" fill="#FFFFFF"/>
  <circle cx="254" cy="170" r="140" fill="#CFEFFF"/>
  <circle cx="1345" cy="158" r="110" fill="#FFF7C7"/>
  <circle cx="1375" cy="760" r="190" fill="#FFD9E8"/>
  <rect x="225" y="330" width="960" height="420" rx="28" fill="#F7FBFE" stroke="#E8EDF2" stroke-width="8"/>
  <rect x="325" y="410" width="760" height="260" rx="22" fill="#FFFFFF"/>
  <rect x="595" y="310" width="220" height="100" rx="18" fill="#CFEFFF"/>
  <rect x="670" y="540" width="110" height="130" rx="18" fill="#FFD9E8"/>
  <rect x="390" y="460" width="95" height="95" rx="16" fill="#CFEFFF"/>
  <rect x="520" y="460" width="95" height="95" rx="16" fill="#FFF7C7"/>
  <rect x="845" y="460" width="95" height="95" rx="16" fill="#FFF7C7"/>
  <rect x="975" y="460" width="95" height="95" rx="16" fill="#CFEFFF"/>
  <rect x="390" y="575" width="95" height="95" rx="16" fill="#FFF7C7"/>
  <rect x="520" y="575" width="95" height="95" rx="16" fill="#CFEFFF"/>
  <rect x="845" y="575" width="95" height="95" rx="16" fill="#CFEFFF"/>
  <rect x="975" y="575" width="95" height="95" rx="16" fill="#FFF7C7"/>
  <path d="M180 750H1230" stroke="#3A3D46" stroke-width="8" stroke-linecap="round" opacity="0.22"/>
  <circle cx="1265" cy="470" r="72" fill="#CFEFFF"/>
  <rect x="1242" y="540" width="48" height="110" rx="24" fill="#3A3D46" opacity="0.8"/>
  <circle cx="1325" cy="470" r="72" fill="#FFD9E8"/>
  <rect x="1302" y="540" width="48" height="110" rx="24" fill="#3A3D46" opacity="0.8"/>
  <circle cx="250" cy="720" r="46" fill="#FFD9E8"/>
  <rect x="220" y="766" width="62" height="110" rx="28" fill="#3A3D46" opacity="0.8"/>
  <rect x="175" y="810" width="55" height="26" rx="13" transform="rotate(-22 175 810)" fill="#3A3D46" opacity="0.8"/>
  <rect x="272" y="810" width="55" height="26" rx="13" transform="rotate(24 272 810)" fill="#3A3D46" opacity="0.8"/>
  <circle cx="395" cy="725" r="46" fill="#FFF7C7"/>
  <rect x="365" y="771" width="62" height="110" rx="28" fill="#3A3D46" opacity="0.8"/>
  <rect x="320" y="814" width="55" height="26" rx="13" transform="rotate(-18 320 814)" fill="#3A3D46" opacity="0.8"/>
  <rect x="416" y="814" width="55" height="26" rx="13" transform="rotate(20 416 814)" fill="#3A3D46" opacity="0.8"/>
  <circle cx="540" cy="720" r="46" fill="#CFEFFF"/>
  <rect x="510" y="766" width="62" height="110" rx="28" fill="#3A3D46" opacity="0.8"/>
  <path d="M520 708C536 722 548 722 564 708" stroke="#3A3D46" stroke-width="8" stroke-linecap="round" opacity="0.75"/>
  <circle cx="1252" cy="698" r="44" fill="#FFF7C7"/>
  <rect x="1225" y="744" width="54" height="100" rx="24" fill="#3A3D46" opacity="0.8"/>
  <path d="M1235 690C1245 678 1258 674 1270 692" stroke="#3A3D46" stroke-width="8" stroke-linecap="round" opacity="0.75"/>
  <line x1="1210" y1="782" x2="1180" y2="810" stroke="#3A3D46" stroke-width="16" stroke-linecap="round" opacity="0.8"/>
  <line x1="1280" y1="782" x2="1315" y2="805" stroke="#3A3D46" stroke-width="16" stroke-linecap="round" opacity="0.8"/>
  <rect x="1340" y="670" width="130" height="160" rx="26" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8"/>
  <line x1="1372" y1="720" x2="1436" y2="720" stroke="#FFD9E8" stroke-width="14" stroke-linecap="round"/>
  <line x1="1372" y1="758" x2="1436" y2="758" stroke="#CFEFFF" stroke-width="14" stroke-linecap="round"/>
  <line x1="1372" y1="796" x2="1412" y2="796" stroke="#FFF7C7" stroke-width="14" stroke-linecap="round"/>
</svg>`,
  "balance-education-wellbeing.svg": `<?xml version="1.0" encoding="UTF-8"?>
<svg width="1200" height="900" viewBox="0 0 1200 900" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="1200" height="900" rx="36" fill="#FFFFFF"/>
  <circle cx="250" cy="220" r="120" fill="#CFEFFF"/>
  <circle cx="948" cy="190" r="96" fill="#FFD9E8"/>
  <circle cx="934" cy="720" r="142" fill="#FFF7C7"/>
  <rect x="560" y="210" width="78" height="390" rx="30" fill="#3A3D46"/>
  <rect x="442" y="580" width="318" height="42" rx="20" fill="#3A3D46"/>
  <path d="M325 330H875" stroke="#3A3D46" stroke-width="22" stroke-linecap="round"/>
  <circle cx="600" cy="320" r="28" fill="#FFF7C7" stroke="#3A3D46" stroke-width="10"/>
  <path d="M360 346L260 528" stroke="#3A3D46" stroke-width="10"/>
  <path d="M840 346L940 528" stroke="#3A3D46" stroke-width="10"/>
  <path d="M280 530H440" stroke="#3A3D46" stroke-width="12" stroke-linecap="round"/>
  <path d="M760 530H920" stroke="#3A3D46" stroke-width="12" stroke-linecap="round"/>
  <path d="M282 530C292 606 344 648 360 648C376 648 428 606 438 530" fill="#CFEFFF" stroke="#3A3D46" stroke-width="10"/>
  <path d="M762 530C772 606 824 648 840 648C856 648 908 606 918 530" fill="#FFD9E8" stroke="#3A3D46" stroke-width="10"/>
  <rect x="314" y="488" width="94" height="58" rx="12" fill="#FFFFFF" stroke="#3A3D46" stroke-width="8"/>
  <line x1="336" y1="507" x2="384" y2="507" stroke="#CFEFFF" stroke-width="10" stroke-linecap="round"/>
  <line x1="336" y1="526" x2="392" y2="526" stroke="#FFF7C7" stroke-width="10" stroke-linecap="round"/>
  <path d="M840 472C820 430 760 440 748 490C736 540 776 578 840 622C904 578 944 540 932 490C920 440 860 430 840 472Z" fill="#FFFFFF" stroke="#3A3D46" stroke-width="8"/>
  <path d="M480 720H720" stroke="#3A3D46" stroke-width="14" stroke-linecap="round" opacity="0.22"/>
</svg>`,
  "stressed-student.svg": `<?xml version="1.0" encoding="UTF-8"?>
<svg width="1250" height="900" viewBox="0 0 1250 900" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="1250" height="900" rx="36" fill="#FFFFFF"/>
  <circle cx="224" cy="176" r="120" fill="#CFEFFF"/>
  <circle cx="1050" cy="180" r="90" fill="#FFF7C7"/>
  <circle cx="1030" cy="720" r="150" fill="#FFD9E8"/>
  <rect x="238" y="585" width="760" height="42" rx="20" fill="#3A3D46" opacity="0.82"/>
  <rect x="360" y="408" width="520" height="195" rx="28" fill="#F7FBFE" stroke="#E8EDF2" stroke-width="8"/>
  <rect x="426" y="324" width="110" height="84" rx="18" fill="#CFEFFF"/>
  <rect x="584" y="328" width="86" height="74" rx="18" fill="#FFD9E8"/>
  <rect x="710" y="340" width="104" height="64" rx="18" fill="#FFF7C7"/>
  <circle cx="624" cy="310" r="74" fill="#FFD9E8"/>
  <path d="M560 300C576 278 595 270 624 270C653 270 672 278 688 300" stroke="#3A3D46" stroke-width="10" stroke-linecap="round"/>
  <path d="M578 354C596 334 610 324 624 324C638 324 652 334 670 354" stroke="#3A3D46" stroke-width="14" stroke-linecap="round"/>
  <rect x="560" y="370" width="128" height="138" rx="30" fill="#3A3D46" opacity="0.82"/>
  <line x1="588" y1="394" x2="530" y2="458" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="658" y1="394" x2="720" y2="458" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <rect x="488" y="465" width="92" height="26" rx="13" transform="rotate(-14 488 465)" fill="#3A3D46" opacity="0.82"/>
  <rect x="668" y="452" width="96" height="26" rx="13" transform="rotate(14 668 452)" fill="#3A3D46" opacity="0.82"/>
  <rect x="300" y="450" width="116" height="86" rx="20" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8" transform="rotate(-10 300 450)"/>
  <line x1="327" y1="480" x2="388" y2="480" stroke="#CFEFFF" stroke-width="12" stroke-linecap="round"/>
  <line x1="327" y1="508" x2="370" y2="508" stroke="#FFD9E8" stroke-width="12" stroke-linecap="round"/>
  <rect x="824" y="454" width="120" height="86" rx="20" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8" transform="rotate(9 824 454)"/>
  <line x1="850" y1="482" x2="916" y2="482" stroke="#FFF7C7" stroke-width="12" stroke-linecap="round"/>
  <line x1="850" y1="510" x2="898" y2="510" stroke="#CFEFFF" stroke-width="12" stroke-linecap="round"/>
  <circle cx="930" cy="320" r="56" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8"/>
  <path d="M930 282V320L958 336" stroke="#3A3D46" stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>
</svg>`,
  "teacher-student-conversation.svg": `<?xml version="1.0" encoding="UTF-8"?>
<svg width="1400" height="900" viewBox="0 0 1400 900" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="1400" height="900" rx="36" fill="#FFFFFF"/>
  <circle cx="235" cy="190" r="118" fill="#CFEFFF"/>
  <circle cx="1160" cy="170" r="96" fill="#FFF7C7"/>
  <circle cx="1120" cy="718" r="150" fill="#FFD9E8"/>
  <rect x="460" y="230" width="470" height="210" rx="28" fill="#F7FBFE" stroke="#E8EDF2" stroke-width="8"/>
  <line x1="530" y1="306" x2="860" y2="306" stroke="#CFEFFF" stroke-width="16" stroke-linecap="round"/>
  <line x1="530" y1="356" x2="778" y2="356" stroke="#FFD9E8" stroke-width="16" stroke-linecap="round"/>
  <circle cx="420" cy="570" r="70" fill="#CFEFFF"/>
  <rect x="380" y="636" width="84" height="152" rx="34" fill="#3A3D46" opacity="0.82"/>
  <line x1="404" y1="678" x2="348" y2="754" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="440" y1="678" x2="518" y2="744" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="402" y1="784" x2="372" y2="842" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="440" y1="784" x2="472" y2="842" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <circle cx="855" cy="592" r="70" fill="#FFD9E8"/>
  <rect x="812" y="658" width="86" height="152" rx="34" fill="#3A3D46" opacity="0.82"/>
  <line x1="838" y1="700" x2="772" y2="760" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="872" y1="700" x2="946" y2="752" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="834" y1="808" x2="802" y2="864" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <line x1="874" y1="808" x2="908" y2="864" stroke="#3A3D46" stroke-width="18" stroke-linecap="round"/>
  <path d="M471 588C536 554 598 546 666 556" stroke="#3A3D46" stroke-width="10" stroke-linecap="round" opacity="0.22"/>
  <path d="M912 602C1002 584 1076 604 1140 654" stroke="#3A3D46" stroke-width="10" stroke-linecap="round" opacity="0.18"/>
  <rect x="995" y="430" width="150" height="102" rx="26" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8"/>
  <circle cx="1043" cy="480" r="10" fill="#CFEFFF"/>
  <circle cx="1072" cy="480" r="10" fill="#FFD9E8"/>
  <circle cx="1101" cy="480" r="10" fill="#FFF7C7"/>
</svg>`,
  "calm-future-school.svg": `<?xml version="1.0" encoding="UTF-8"?>
<svg width="1600" height="960" viewBox="0 0 1600 960" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect width="1600" height="960" rx="40" fill="#FFFFFF"/>
  <circle cx="280" cy="190" r="140" fill="#CFEFFF"/>
  <circle cx="1330" cy="220" r="120" fill="#FFF7C7"/>
  <circle cx="1300" cy="760" r="200" fill="#FFD9E8"/>
  <circle cx="805" cy="430" r="150" fill="#FFF7C7" opacity="0.9"/>
  <path d="M150 720C286 640 426 628 570 680C706 730 854 734 990 674C1114 620 1248 614 1450 706V814H150V720Z" fill="#F7FBFE"/>
  <rect x="560" y="464" width="486" height="218" rx="28" fill="#FFFFFF" stroke="#E8EDF2" stroke-width="8"/>
  <rect x="655" y="518" width="296" height="126" rx="18" fill="#CFEFFF"/>
  <rect x="765" y="424" width="74" height="96" rx="18" fill="#FFD9E8"/>
  <rect x="700" y="556" width="54" height="54" rx="12" fill="#FFFFFF"/>
  <rect x="773" y="556" width="54" height="54" rx="12" fill="#FFF7C7"/>
  <rect x="846" y="556" width="54" height="54" rx="12" fill="#FFFFFF"/>
  <path d="M515 818H1095" stroke="#3A3D46" stroke-width="12" stroke-linecap="round" opacity="0.18"/>
  <path d="M268 770C318 700 386 684 456 720" stroke="#CFEFFF" stroke-width="24" stroke-linecap="round" opacity="0.9"/>
  <path d="M1148 770C1194 706 1262 686 1334 724" stroke="#FFD9E8" stroke-width="24" stroke-linecap="round" opacity="0.9"/>
</svg>`,
};

const slideDefinitions = [
  {
    number: 1,
    transition: { type: "fade" },
    animationTargets: [
      "Wie Schule wirkt:",
      "Seminarfachkurs Glück",
    ],
    notes:
      "Ich möchte heute zeigen, wie Schule nicht nur Leistungen, sondern auch das Wohlbefinden von Schülerinnen und Schülern prägt. Ausgangspunkt meiner Facharbeit war die Beobachtung, dass Schule einerseits Wissen vermittelt und Zukunftschancen eröffnet, andererseits aber auch Druck, Vergleiche und Unsicherheit erzeugen kann. Gerade im Seminarfachkurs Glück ist das spannend, weil Wohlbefinden nicht nur von individuellen Eigenschaften abhängt, sondern auch von den Strukturen, in denen Jugendliche lernen. Deshalb richte ich den Blick heute auf das Schulsystem selbst. Im Zentrum steht die Frage, welche Wirkungen Schulformen, Bewertungen und schulische Rahmenbedingungen auf das Erleben von Schule haben. Warum dieses Thema so relevant ist, zeige ich auf der nächsten Folie.",
  },
  {
    number: 2,
    transition: { type: "zoom" },
    animationTargets: [
      "Schule prägt den Alltag",
      "Inwiefern beeinflussen strukturelle Unterschiede",
    ],
    notes:
      "Das Thema ist wichtig, weil Schule ein zentraler Lebensraum für Kinder und Jugendliche ist. Ein großer Teil des Alltags spielt sich dort ab, und schulische Erfahrungen wirken auf Selbstbild, Motivation und Zukunftsperspektiven. In der Facharbeit wird außerdem deutlich, dass die öffentliche Diskussion über Leistungsdruck und psychische Belastungen zunimmt. Bildung ist zugleich entscheidend für soziale Teilhabe, berufliche Chancen und demokratische Stabilität. Deshalb reicht es nicht, nur auf Leistungen zu schauen. Wenn Bildung erfolgreich sein soll, muss auch das Wohlbefinden mitgedacht werden. Aus genau diesem Spannungsfeld ergibt sich die Leitfrage meiner Arbeit, die ich unten hervorgehoben habe und die den roten Faden für die gesamte Präsentation bildet.",
  },
  {
    number: 3,
    transition: { type: "wipe" },
    animationTargets: [
      "Strukturmerkmale",
      "Wohlbefinden",
      "Bildungsgerechtigkeit",
    ],
    notes:
      "Bevor ich auf das deutsche Schulsystem eingehe, müssen drei Grundbegriffe geklärt werden. Erstens meint Strukturmerkmale die institutionellen Rahmenbedingungen von Schule, also zum Beispiel Schulformen, Notensystem, Versetzungsordnungen und Abschlüsse. Zweitens umfasst Wohlbefinden laut Facharbeit nicht nur gute Laune, sondern psychische Gesundheit, soziale Beziehungen und Lebenszufriedenheit. Es geht also darum, wie sicher, eingebunden und belastet sich Jugendliche in Schule fühlen. Drittens ist Bildungsgerechtigkeit die Frage, ob Bildungszugang und Bildungschancen fair verteilt sind und nicht zu stark von sozialer Herkunft abhängen. Diese drei Begriffe gehören zusammen und bilden die Grundlage für alle späteren Bewertungen. Mit diesem Begriffsrahmen können wir jetzt auf die Struktur des deutschen Systems schauen.",
  },
  {
    number: 4,
    transition: { type: "cover" },
    animationTargets: [
      "Gymnasium",
      "Realschule",
      "Hauptschule",
      "Gesamtschule",
      "Berufliche Schulen",
    ],
    notes:
      "Hier sieht man die wichtigsten Schulformen, die in der Arbeit miteinander verglichen werden. Das Gymnasium ist am stärksten akademisch ausgerichtet und führt zum Abitur. Die Realschule nimmt eine mittlere Position ein und verbindet Allgemeinbildung mit mehr Praxisbezug. Die Hauptschule war traditionell auf praktische und berufliche Wege ausgerichtet, wird aber heute auch wegen Stigmatisierung und sozialer Segregation kritisch gesehen. Die Gesamtschule soll längeres gemeinsames Lernen ermöglichen und verschiedene Abschlüsse unter einem Dach verbinden. Berufliche Schulen sind schließlich eng mit dem dualen Ausbildungssystem verbunden und gelten international als besondere Stärke Deutschlands. Mit diesen Unterschieden im Hinterkopf wird deutlicher, warum Schulstruktur nicht neutral ist, sondern Erfahrungen und Anforderungen verändert.",
  },
  {
    number: 5,
    transition: { type: "circle" },
    animationTargets: [
      "Noten",
      "Stress",
      "Leistungsdruck kann Motivation fördern",
    ],
    notes:
      "Ein zentrales Ergebnis der Facharbeit ist, dass Leistungsdruck das Wohlbefinden stark beeinflussen kann. Dieser Druck entsteht vor allem durch Noten, Prüfungen, Konkurrenz und Erwartungen von Eltern, Lehrkräften oder Gesellschaft. Solche Anforderungen müssen nicht grundsätzlich negativ sein. Für manche Schülerinnen und Schüler wirken sie sogar motivierend. Problematisch wird es dann, wenn aus Leistungsanforderungen dauerhafter Stress wird. In der Arbeit werden als mögliche Folgen zum Beispiel Schlafprobleme, Unsicherheit, psychosomatische Beschwerden und ein geringeres Selbstwertgefühl beschrieben. Entscheidend ist also nicht nur, dass Leistung bewertet wird, sondern wie stark Erfolg und Misserfolg mit dem eigenen Wert verbunden werden. Nach dem Leistungsdruck kommt nun ein zweiter wichtiger Faktor: das Schulklima.",
  },
  {
    number: 6,
    transition: { type: "push" },
    animationTargets: [
      "Lehrer-Schüler-Verhältnis",
      "Mitbestimmung",
      "Feedbackkultur",
    ],
    notes:
      "Neben Leistung wirkt vor allem das soziale Klima der Schule auf das Wohlbefinden. Die Facharbeit nennt dabei drei Bereiche. Erstens ist das Lehrer-Schüler-Verhältnis wichtig: Wer sich ernst genommen fühlt und Unterstützung erlebt, empfindet Schule meist als weniger belastend. Zweitens spielt Mitbestimmung eine Rolle. Wenn Jugendliche an Entscheidungen beteiligt werden, stärkt das ihre Selbstwirksamkeit und das Gefühl, dass ihre Meinung zählt. Drittens kommt es auf die Feedbackkultur an. Qualitative Rückmeldungen können Lernprozesse begleiten und individuelle Entwicklung sichtbarer machen als reine Noten. Zusammengefasst zeigt sich: Schule wird dann eher als förderlicher Ort erlebt, wenn Beziehungen, Beteiligung und Feedback stimmen. Genau deshalb lohnt sich jetzt ein kurzer Blick auf eure eigene Wahrnehmung.",
  },
  {
    number: 7,
    transition: { type: "dissolve" },
    animationTargets: [
      "Was beeinflusst euer Wohlbefinden",
      "Alle Faktoren spielen laut Facharbeit",
    ],
    notes:
      "An dieser Stelle würde ich das Publikum kurz einbinden. Die Frage lautet: Was beeinflusst euer Wohlbefinden in der Schule am stärksten? Die vier Antwortmöglichkeiten greifen zentrale Punkte aus der Facharbeit auf: Notendruck, Lehrkräfte, Mitschüler und Zeitstress. Ein kurzes Handzeichen reicht hier vollkommen aus. Wichtig ist für meine Auswertung vor allem die Erkenntnis, dass diese Faktoren nicht isoliert auftreten. Die Facharbeit zeigt vielmehr, dass Wohlbefinden immer aus einem Zusammenspiel von schulischen Anforderungen, sozialen Beziehungen und institutionellen Rahmenbedingungen entsteht. Diese kleine Umfrage schafft also einen guten Übergang zur nächsten Folie, auf der ich zeige, dass andere Bildungssysteme manche dieser Faktoren anders gestalten.",
  },
  {
    number: 8,
    transition: { type: "plus" },
    animationTargets: [
      "Deutschland",
      "Finnland",
      "Angelsächsische Systeme",
      "Kein System ist perfekt.",
    ],
    notes:
      "Der internationale Vergleich hilft dabei, das deutsche System besser einzuordnen. In Deutschland fallen vor allem die frühe Aufteilung und die starke Bedeutung von Noten und Prüfungen auf. Finnland steht in der Arbeit beispielhaft für ein Modell mit längerem gemeinsamen Lernen und späterer Leistungsdifferenzierung. In den frühen Schuljahren stehen dort individuelle Lernfortschritte und qualitative Rückmeldungen stärker im Vordergrund. Angelsächsische Systeme werden wiederum mit mehr Wahlmöglichkeiten, projektorientiertem Lernen und größerer Praxisnähe verbunden. Gleichzeitig macht die Facharbeit klar, dass es kein perfektes Bildungssystem gibt. Jedes System hat eigene Stärken und Herausforderungen. Genau deshalb ist es sinnvoll, den deutschen Fall nun ausgewogen nach Stärken und Schwächen zu bewerten.",
  },
  {
    number: 9,
    transition: { type: "split" },
    animationTargets: [
      "Stärken",
      "Schwächen",
    ],
    notes:
      "Die Facharbeit bewertet das deutsche Schulsystem bewusst differenziert. Zu den Stärken gehören die verschiedenen Bildungswege, das duale Ausbildungssystem und die hohen akademischen Standards, besonders im Gymnasium. Diese Strukturen können gezielte Förderung ermöglichen und den Übergang in Studium oder Beruf gut vorbereiten. Gleichzeitig gibt es klare Schwächen. Besonders kritisch sind die frühe Selektion, der Leistungsdruck und die fortbestehenden sozialen Ungleichheiten. Die Arbeit zeigt, dass gerade die frühe Aufteilung nach der Grundschule Bildungswege sehr früh festlegt und dass der Bildungserfolg in Deutschland vergleichsweise stark mit sozialer Herkunft zusammenhängt. Aus dieser doppelten Perspektive ergeben sich auch die Reformansätze, die ich als Nächstes vorstelle.",
  },
  {
    number: 10,
    transition: { type: "wheel" },
    animationTargets: [
      "Mehr Lebenskompetenzen",
      "Spätere Leistungsdifferenzierung",
      "Alternative Bewertungsformen",
    ],
    notes:
      "Aus den beschriebenen Problemen leitet die Facharbeit drei zentrale Reformansätze ab. Erstens soll Schule mehr Lebenskompetenzen vermitteln, also zum Beispiel finanzielle Grundbildung, Medienkompetenz und alltagsrelevante Fähigkeiten. Zweitens wird eine spätere Leistungsdifferenzierung diskutiert, damit Jugendliche mehr Zeit haben, ihre Stärken zu entwickeln und fundiertere Bildungsentscheidungen zu treffen. Drittens geht es um alternative Bewertungsformen. Ausführliche Lernfeedbacks, Portfolioarbeit oder projektbasierte Bewertung könnten Lernprozesse stärker in den Mittelpunkt rücken als reine Prüfungsergebnisse. Diese Reformen würden das System nicht vollständig neu erfinden, aber sie könnten Bildung und Wohlbefinden besser miteinander verbinden. Damit komme ich zum abschließenden Fazit.",
  },
  {
    number: 11,
    transition: { type: "fade" },
    animationTargets: [
      "Schule beeinflusst mehr als nur Noten.",
      "Strukturen wirken auf das Wohlbefinden.",
      "Danke für eure Aufmerksamkeit!",
    ],
    notes:
      "Im Fazit lässt sich die Leitfrage klar beantworten: Strukturelle Unterschiede im deutschen Schulsystem beeinflussen das Wohlbefinden von Schülerinnen und Schülern in erheblichem Maße. Entscheidend sind dabei nicht nur Unterrichtsinhalte, sondern vor allem die Schulstruktur, Leistungsbewertung und das soziale Klima. Leistungsdruck und Konkurrenz können motivierend sein, aber ebenso Stress und Unsicherheit auslösen. Gleichzeitig zeigt die Arbeit, dass ein positives Schulklima, Mitbestimmung und gute Feedbackkultur entlastend wirken können. Der Blick ins Ausland macht außerdem deutlich, dass andere Wege möglich sind. Schule beeinflusst also mehr als nur Noten - sie wirkt auf Chancen, Selbstbild und Lebenszufriedenheit. Damit bedanke ich mich für eure Aufmerksamkeit und freue mich auf Fragen oder Rückmeldungen.",
  },
];

function ensureDirectories() {
  fs.mkdirSync(ASSET_DIR, { recursive: true });
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function writeAssets() {
  for (const [name, markup] of Object.entries(assets)) {
    fs.writeFileSync(path.join(ASSET_DIR, name), markup, "utf8");
  }
}

function writeNotesMarkdown() {
  const sections = slideDefinitions
    .map(
      (slide) =>
        `## Folie ${slide.number}\n\n${slide.notes}\n`,
    )
    .join("\n");
  fs.writeFileSync(
    path.join(OUTPUT_DIR, "Moderationsnotizen.md"),
    `# Moderationsnotizen\n\n${sections}`,
    "utf8",
  );
}

function writeMetadata() {
  fs.writeFileSync(
    path.join(OUTPUT_DIR, "presentation-metadata.json"),
    JSON.stringify(slideDefinitions.map(({ number, transition, animationTargets }) => ({
      number,
      transition,
      animationTargets,
    })), null, 2),
    "utf8",
  );
}

function addShadowCard(slide, x, y, w, h, fillColor = palette.white, radius = 0.22) {
  slide.addShape("roundRect", {
    x: x + 0.05,
    y: y + 0.06,
    w,
    h,
    rectRadius: radius,
    line: { color: palette.shadow, transparency: 100 },
    fill: { color: palette.ink, transparency: 92 },
  });
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h,
    rectRadius: radius,
    line: { color: palette.line, pt: 1 },
    fill: { color: fillColor },
  });
}

function addSoftBackground(slide, variant = 0) {
  slide.background = { color: palette.white };
  const variants = [
    [
      { x: -0.6, y: -0.35, w: 2.6, h: 2.0, color: palette.blue, transparency: 10 },
      { x: 10.6, y: -0.2, w: 2.0, h: 1.7, color: palette.yellow, transparency: 6 },
      { x: 10.5, y: 5.4, w: 2.7, h: 2.2, color: palette.pink, transparency: 8 },
    ],
    [
      { x: -0.8, y: 5.25, w: 2.8, h: 2.1, color: palette.pink, transparency: 12 },
      { x: 10.7, y: -0.25, w: 1.8, h: 1.4, color: palette.blue, transparency: 10 },
      { x: 8.8, y: 5.6, w: 2.1, h: 1.6, color: palette.yellow, transparency: 8 },
    ],
    [
      { x: -0.55, y: -0.35, w: 2.2, h: 1.8, color: palette.yellow, transparency: 8 },
      { x: 10.6, y: 0.05, w: 2.3, h: 1.7, color: palette.pink, transparency: 10 },
      { x: 9.8, y: 5.45, w: 2.9, h: 2.1, color: palette.blue, transparency: 10 },
    ],
  ];
  for (const blob of variants[variant % variants.length]) {
    slide.addShape("ellipse", {
      x: blob.x,
      y: blob.y,
      w: blob.w,
      h: blob.h,
      line: { color: blob.color, transparency: 100 },
      fill: { color: blob.color, transparency: blob.transparency },
    });
  }
}

function addEyebrow(slide, text, x = 0.7, y = 0.45, w = 3.2) {
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h: 0.42,
    rectRadius: 0.2,
    line: { color: palette.line, pt: 0.8 },
    fill: { color: palette.white },
  });
  slide.addText(text, {
    x: x + 0.16,
    y: y + 0.08,
    w: w - 0.3,
    h: 0.22,
    fontFace: "Aptos",
    fontSize: 11,
    color: palette.muted,
    bold: true,
    margin: 0,
  });
}

function addTitle(slide, title, x = 0.7, y = 1.0, w = 6.2, h = 0.9, fontSize = 24) {
  slide.addText(title, {
    x,
    y,
    w,
    h,
    fontFace: "Aptos Display",
    fontSize,
    color: palette.ink,
    bold: true,
    margin: 0,
    fit: "shrink",
  });
}

function addBodyText(slide, text, opts) {
  slide.addText(text, {
    fontFace: "Aptos",
    fontSize: 15,
    color: palette.ink,
    margin: 0,
    ...opts,
  });
}

function bulletRuns(items, fontSize = 18) {
  return items.map((item, index) => ({
    text: item,
    options: {
      breakLine: index > 0,
      bullet: { indent: 18 },
      color: palette.ink,
      fontSize,
      paraSpaceAfterPt: 12,
    },
  }));
}

function lineRuns(items, fontSize = 16, color = palette.ink, boldFirst = false) {
  return items.map((item, index) => ({
    text: item,
    options: {
      breakLine: index > 0,
      color,
      fontSize,
      bold: boldFirst && index === 0,
      paraSpaceAfterPt: 12,
    },
  }));
}

function addImagePanel(slide, imagePath, x, y, w, h, fillColor = palette.slate) {
  addShadowCard(slide, x, y, w, h, palette.white);
  slide.addShape("roundRect", {
    x: x + 0.18,
    y: y + 0.18,
    w: w - 0.36,
    h: h - 0.36,
    rectRadius: 0.18,
    line: { color: fillColor, transparency: 100 },
    fill: { color: fillColor, transparency: 10 },
  });
  slide.addImage({
    path: imagePath,
    x: x + 0.18,
    y: y + 0.18,
    w: w - 0.36,
    h: h - 0.36,
  });
}

function addQuoteBox(slide, text, x, y, w, h, fillColor = palette.yellow) {
  addShadowCard(slide, x, y, w, h, fillColor);
  slide.addText(text, {
    x: x + 0.24,
    y: y + 0.16,
    w: w - 0.48,
    h: h - 0.32,
    fontFace: "Aptos",
    fontSize: 17,
    color: palette.ink,
    bold: true,
    italic: true,
    align: "center",
    valign: "mid",
    margin: 0,
    fit: "shrink",
  });
}

function addMiniLabel(slide, text, x, y, w, fillColor) {
  slide.addShape("roundRect", {
    x,
    y,
    w,
    h: 0.36,
    rectRadius: 0.18,
    line: { color: fillColor, transparency: 100 },
    fill: { color: fillColor },
  });
  slide.addText(text, {
    x: x + 0.12,
    y: y + 0.07,
    w: w - 0.24,
    h: 0.2,
    fontFace: "Aptos",
    fontSize: 10.5,
    color: palette.ink,
    bold: true,
    margin: 0,
    align: "center",
  });
}

function addCard(slide, { x, y, w, h, title, body, fillColor = palette.white, label }) {
  addShadowCard(slide, x, y, w, h, fillColor);
  if (label) {
    addMiniLabel(slide, label, x + 0.18, y + 0.14, Math.min(1.35, w - 0.36), palette.white);
  }
  slide.addText(title, {
    x: x + 0.18,
    y: y + (label ? 0.5 : 0.24),
    w: w - 0.36,
    h: 0.45,
    fontFace: "Aptos Display",
    fontSize: 17,
    color: palette.ink,
    bold: true,
    margin: 0,
    fit: "shrink",
  });
  slide.addText(body, {
    x: x + 0.18,
    y: y + (label ? 0.98 : 0.72),
    w: w - 0.36,
    h: h - (label ? 1.12 : 0.88),
    fontFace: "Aptos",
    fontSize: 13.5,
    color: palette.muted,
    margin: 0,
    fit: "shrink",
    breakLine: false,
  });
}

function registerSlideMeta(number, transition, animationTargets) {
  slidesMeta.push({ number, transition, animationTargets });
}

function buildSlide1(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 0);
  addEyebrow(slide, "Seminarfachkurs Glück", 0.7, 0.52, 2.3);
  slide.addText("Wie Schule wirkt:", {
    x: 0.7,
    y: 1.25,
    w: 5.7,
    h: 0.7,
    fontFace: "Aptos Display",
    fontSize: 24,
    bold: true,
    color: palette.ink,
    margin: 0,
  });
  slide.addText(
    "Strukturelle Unterschiede\nund ihr Effekt auf das Wohlbefinden\nvon Schülerinnen und Schülern",
    {
      x: 0.7,
      y: 1.85,
      w: 5.9,
      h: 1.8,
      fontFace: "Aptos Display",
      fontSize: 24,
      color: palette.ink,
      bold: true,
      margin: 0,
      fit: "shrink",
    },
  );
  slide.addText("Seminarfachkurs Glück", {
    x: 0.7,
    y: 4.05,
    w: 3.3,
    h: 0.35,
    fontFace: "Aptos",
    fontSize: 14,
    color: palette.muted,
    margin: 0,
  });
  addImagePanel(
    slide,
    path.join(ASSET_DIR, "title-school-illustration.svg"),
    7.0,
    0.9,
    5.55,
    5.7,
    palette.blue,
  );
  slide.addText("Schule als Lebensraum zwischen Chance, Druck und Wohlbefinden", {
    x: 0.7,
    y: 5.0,
    w: 5.4,
    h: 0.6,
    fontFace: "Aptos",
    fontSize: 16,
    color: palette.ink,
    bold: true,
    margin: 0,
    fit: "shrink",
  });
  slide.addShape("line", {
    x: 0.7,
    y: 5.65,
    w: 4.6,
    h: 0,
    line: { color: palette.line, pt: 1.5 },
  });
  slide.addNotes(slideDefinitions[0].notes);
  registerSlideMeta(1, slideDefinitions[0].transition, slideDefinitions[0].animationTargets);
}

function buildSlide2(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 1);
  addEyebrow(slide, "Relevanz des Themas");
  addTitle(slide, "Warum ist das Thema wichtig?", 0.7, 1.0, 5.6, 0.6, 24);
  slide.addText(bulletRuns([
    "Schule prägt den Alltag von Kindern und Jugendlichen.",
    "Leistungsdruck und psychische Belastungen werden intensiver diskutiert.",
    "Bildung beeinflusst Teilhabe, Berufschancen und Zukunftsperspektiven.",
    "Erfolgreiche Bildung braucht auch Wohlbefinden.",
  ], 18), {
    x: 0.9,
    y: 1.95,
    w: 5.4,
    h: 2.8,
    margin: 0,
    valign: "top",
  });
  addImagePanel(
    slide,
    path.join(ASSET_DIR, "balance-education-wellbeing.svg"),
    7.1,
    1.2,
    5.05,
    3.7,
    palette.yellow,
  );
  addQuoteBox(
    slide,
    "Inwiefern beeinflussen strukturelle Unterschiede im deutschen Schulsystem das Wohlbefinden von Schülerinnen und Schülern?",
    0.72,
    5.15,
    11.85,
    1.15,
    palette.blue,
  );
  slide.addNotes(slideDefinitions[1].notes);
  registerSlideMeta(2, slideDefinitions[1].transition, slideDefinitions[1].animationTargets);
}

function buildSlide3(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 2);
  addEyebrow(slide, "Theoretische Grundlagen");
  addTitle(slide, "Grundlagen", 0.7, 1.0, 3.5, 0.6, 24);

  slide.addShape("ellipse", {
    x: 5.12,
    y: 2.38,
    w: 2.0,
    h: 1.25,
    line: { color: palette.line, pt: 1 },
    fill: { color: palette.white },
  });
  slide.addText("Schule", {
    x: 5.42,
    y: 2.75,
    w: 1.4,
    h: 0.3,
    fontFace: "Aptos Display",
    fontSize: 19,
    color: palette.ink,
    bold: true,
    margin: 0,
    align: "center",
  });

  slide.addShape("line", {
    x: 4.15,
    y: 2.95,
    w: 1.15,
    h: -0.75,
    line: { color: palette.line, pt: 1.3 },
  });
  slide.addShape("line", {
    x: 7.0,
    y: 2.95,
    w: 1.2,
    h: -0.75,
    line: { color: palette.line, pt: 1.3 },
  });
  slide.addShape("line", {
    x: 6.12,
    y: 3.62,
    w: 0,
    h: 1.0,
    line: { color: palette.line, pt: 1.3 },
  });

  addCard(slide, {
    x: 0.85,
    y: 1.72,
    w: 3.25,
    h: 2.0,
    title: "Strukturmerkmale",
    body: "Schulformen\nNotensystem\nVersetzungen\nAbschlüsse",
    fillColor: palette.blue,
  });
  addCard(slide, {
    x: 8.25,
    y: 1.72,
    w: 3.25,
    h: 2.0,
    title: "Wohlbefinden",
    body: "psychische Gesundheit\nsoziale Beziehungen\nLebenszufriedenheit",
    fillColor: palette.pink,
  });
  addCard(slide, {
    x: 4.05,
    y: 4.25,
    w: 4.15,
    h: 1.7,
    title: "Bildungsgerechtigkeit",
    body: "faire Chancen unabhängig von sozialer Herkunft",
    fillColor: palette.yellow,
  });
  slide.addNotes(slideDefinitions[2].notes);
  registerSlideMeta(3, slideDefinitions[2].transition, slideDefinitions[2].animationTargets);
}

function buildSlide4(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 0);
  addEyebrow(slide, "Strukturelle Unterschiede");
  addTitle(slide, "Das deutsche Schulsystem", 0.7, 1.0, 5.2, 0.6, 24);
  slide.addText("Fünf Schulformen im knappen Vergleich", {
    x: 0.7,
    y: 1.52,
    w: 4.2,
    h: 0.28,
    fontFace: "Aptos",
    fontSize: 13,
    color: palette.muted,
    margin: 0,
  });

  const cards = [
    ["🎓 Gymnasium", "Abitur\nhohe akademische Anforderungen", palette.blue],
    ["🧭 Realschule", "mittlerer Abschluss\nAllgemeinbildung und Praxis", palette.white],
    ["🛠 Hauptschule", "praktisch ausgerichtet\nRisiko von Stigmatisierung", palette.yellow],
    ["🏫 Gesamtschule", "länger gemeinsam lernen\nmehrere Abschlüsse möglich", palette.pink],
    ["💼 Berufliche Schulen", "duales System\nSchule und Betrieb verbunden", palette.white],
  ];

  cards.forEach(([title, body, fillColor], idx) => {
    addCard(slide, {
      x: 0.85 + (idx % 2) * 4.15 + (idx === 4 ? 2.08 : 0),
      y: 1.95 + Math.floor(idx / 2) * 1.65 + (idx === 4 ? 1.6 : 0),
      w: 3.7,
      h: 1.35,
      title,
      body,
      fillColor,
    });
  });

  slide.addShape("line", {
    x: 6.55,
    y: 2.62,
    w: 0,
    h: 2.78,
    line: { color: palette.line, pt: 1.3 },
  });
  slide.addNotes(slideDefinitions[3].notes);
  registerSlideMeta(4, slideDefinitions[3].transition, slideDefinitions[3].animationTargets);
}

function buildSlide5(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 1);
  addEyebrow(slide, "Schule und Wohlbefinden");
  addTitle(slide, "Leistungsdruck und Wohlbefinden", 0.7, 1.0, 6.2, 0.65, 24);
  addImagePanel(
    slide,
    path.join(ASSET_DIR, "stressed-student.svg"),
    8.3,
    1.0,
    4.1,
    3.05,
    palette.pink,
  );

  addShadowCard(slide, 0.85, 1.95, 3.2, 2.75, palette.white);
  addMiniLabel(slide, "Leistungsdruck", 1.05, 2.15, 1.7, palette.blue);
  slide.addText(lineRuns(["Noten", "Prüfungen", "Konkurrenz", "Erwartungen"], 17), {
    x: 1.1,
    y: 2.72,
    w: 2.6,
    h: 1.45,
    margin: 0,
  });

  addShadowCard(slide, 4.35, 1.95, 3.2, 2.75, palette.white);
  addMiniLabel(slide, "Folgen", 4.55, 2.15, 1.0, palette.pink);
  slide.addText(lineRuns(["Stress", "Schlafprobleme", "Unsicherheit", "geringeres Selbstwertgefühl"], 17), {
    x: 4.6,
    y: 2.72,
    w: 2.55,
    h: 1.6,
    margin: 0,
    fit: "shrink",
  });

  addQuoteBox(
    slide,
    "Leistungsdruck kann Motivation fördern, aber auch das Wohlbefinden beeinträchtigen.",
    0.85,
    5.1,
    11.55,
    0.95,
    palette.yellow,
  );
  slide.addNotes(slideDefinitions[4].notes);
  registerSlideMeta(5, slideDefinitions[4].transition, slideDefinitions[4].animationTargets);
}

function buildSlide6(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 2);
  addEyebrow(slide, "Beziehungsebene");
  addTitle(slide, "Schulklima", 0.7, 1.0, 3.0, 0.6, 24);
  addImagePanel(
    slide,
    path.join(ASSET_DIR, "teacher-student-conversation.svg"),
    8.0,
    0.95,
    4.4,
    2.9,
    palette.blue,
  );

  addCard(slide, {
    x: 0.85,
    y: 2.0,
    w: 3.55,
    h: 2.8,
    title: "Lehrer-Schüler-Verhältnis",
    body: "Unterstützung und Respekt können Stress reduzieren.\nPositive Beziehungen fördern Motivation und Orientierung.",
    fillColor: palette.blue,
  });
  addCard(slide, {
    x: 4.75,
    y: 2.0,
    w: 3.55,
    h: 2.8,
    title: "Mitbestimmung",
    body: "Beteiligung am Schulalltag stärkt Selbstwirksamkeit.\nSchule wird eher als Ort erlebt, an dem die eigene Meinung zählt.",
    fillColor: palette.yellow,
  });
  addCard(slide, {
    x: 8.65,
    y: 4.1,
    w: 3.55,
    h: 1.85,
    title: "Feedbackkultur",
    body: "Qualitative Rückmeldungen begleiten Lernprozesse und machen Entwicklung sichtbarer.",
    fillColor: palette.pink,
  });
  slide.addNotes(slideDefinitions[5].notes);
  registerSlideMeta(6, slideDefinitions[5].transition, slideDefinitions[5].animationTargets);
}

function buildSlide7(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 0);
  addEyebrow(slide, "Publikum einbinden");
  addTitle(slide, "Kurze Umfrage", 0.7, 1.0, 3.5, 0.6, 24);
  addShadowCard(slide, 1.0, 1.65, 11.15, 2.0, palette.white);
  slide.addText("Was beeinflusst euer Wohlbefinden in der Schule am stärksten?", {
    x: 1.55,
    y: 2.22,
    w: 10.0,
    h: 0.8,
    fontFace: "Aptos Display",
    fontSize: 25,
    color: palette.ink,
    bold: true,
    margin: 0,
    align: "center",
    fit: "shrink",
  });

  const buttons = [
    ["📚 Notendruck", 1.05, 4.2, palette.blue],
    ["👩‍🏫 Lehrkräfte", 4.15, 4.2, palette.pink],
    ["👥 Mitschüler", 7.25, 4.2, palette.yellow],
    ["🕒 Zeitstress", 10.35, 4.2, palette.blue],
  ];
  buttons.forEach(([label, x, y, fillColor]) => {
    slide.addShape("roundRect", {
      x,
      y,
      w: 2.0,
      h: 0.72,
      rectRadius: 0.24,
      line: { color: fillColor, transparency: 100 },
      fill: { color: fillColor },
    });
    slide.addText(label, {
      x: x + 0.1,
      y: y + 0.19,
      w: 1.8,
      h: 0.25,
      fontFace: "Aptos",
      fontSize: 12,
      color: palette.ink,
      bold: true,
      margin: 0,
      align: "center",
      fit: "shrink",
    });
  });

  slide.addText("Kurzes Handzeichen", {
    x: 4.65,
    y: 5.2,
    w: 4.0,
    h: 0.28,
    fontFace: "Aptos",
    fontSize: 14,
    color: palette.muted,
    bold: true,
    align: "center",
    margin: 0,
  });
  addQuoteBox(
    slide,
    "Alle Faktoren spielen laut Facharbeit eine wichtige Rolle.",
    2.2,
    5.65,
    8.95,
    0.76,
    palette.pink,
  );
  slide.addNotes(slideDefinitions[6].notes);
  registerSlideMeta(7, slideDefinitions[6].transition, slideDefinitions[6].animationTargets);
}

function buildSlide8(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 1);
  addEyebrow(slide, "Internationaler Vergleich");
  addTitle(slide, "Deutschland - Finnland - angelsächsische Systeme", 0.7, 1.0, 8.8, 0.6, 22);

  addCard(slide, {
    x: 0.85,
    y: 2.0,
    w: 3.55,
    h: 2.7,
    title: "Deutschland",
    body: "frühe Aufteilung\nNoten und Prüfungen zentral",
    fillColor: palette.blue,
  });
  addCard(slide, {
    x: 4.9,
    y: 1.7,
    w: 3.55,
    h: 3.0,
    title: "Finnland",
    body: "längeres gemeinsames Lernen\nspätere Leistungsdifferenzierung\nindividuelle Lernfortschritte stärker im Blick",
    fillColor: palette.yellow,
  });
  addCard(slide, {
    x: 8.95,
    y: 2.0,
    w: 3.55,
    h: 2.7,
    title: "Angelsächsische Systeme",
    body: "mehr Praxisnähe\nmehr Wahlmöglichkeiten\nhäufig projektorientiertes Lernen",
    fillColor: palette.pink,
  });
  slide.addShape("chevron", {
    x: 5.8,
    y: 4.95,
    w: 1.8,
    h: 0.55,
    line: { color: palette.line, transparency: 100 },
    fill: { color: palette.blue, transparency: 20 },
  });
  addQuoteBox(slide, "Kein System ist perfekt.", 4.05, 5.55, 5.25, 0.76, palette.white);
  slide.addNotes(slideDefinitions[7].notes);
  registerSlideMeta(8, slideDefinitions[7].transition, slideDefinitions[7].animationTargets);
}

function buildSlide9(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 2);
  addEyebrow(slide, "Kritische Bewertung");
  addTitle(slide, "Stärken und Schwächen", 0.7, 1.0, 5.3, 0.6, 24);
  addShadowCard(slide, 0.9, 1.85, 5.55, 3.9, palette.blue);
  addShadowCard(slide, 6.88, 1.85, 5.55, 3.9, palette.pink);

  slide.addText("Stärken", {
    x: 1.18,
    y: 2.15,
    w: 2.3,
    h: 0.35,
    fontFace: "Aptos Display",
    fontSize: 22,
    bold: true,
    color: palette.ink,
    margin: 0,
  });
  slide.addText(lineRuns([
    "verschiedene Bildungswege",
    "duales Ausbildungssystem",
    "hohe akademische Standards",
  ], 18), {
    x: 1.18,
    y: 2.8,
    w: 4.65,
    h: 1.55,
    margin: 0,
  });

  slide.addText("Schwächen", {
    x: 7.16,
    y: 2.15,
    w: 2.5,
    h: 0.35,
    fontFace: "Aptos Display",
    fontSize: 22,
    bold: true,
    color: palette.ink,
    margin: 0,
  });
  slide.addText(lineRuns([
    "frühe Selektion",
    "Leistungsdruck",
    "soziale Ungleichheiten",
  ], 18), {
    x: 7.16,
    y: 2.8,
    w: 4.7,
    h: 1.55,
    margin: 0,
  });

  slide.addText("Das deutsche System verbindet Qualität und Differenzierung - aber nicht für alle gleich entlastend.", {
    x: 1.18,
    y: 5.05,
    w: 10.8,
    h: 0.36,
    fontFace: "Aptos",
    fontSize: 14,
    color: palette.muted,
    bold: true,
    margin: 0,
    align: "center",
  });
  slide.addNotes(slideDefinitions[8].notes);
  registerSlideMeta(9, slideDefinitions[8].transition, slideDefinitions[8].animationTargets);
}

function buildSlide10(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 0);
  addEyebrow(slide, "Reformansätze");
  addTitle(slide, "Pfeil in Richtung Zukunft", 0.7, 1.0, 5.0, 0.6, 24);

  slide.addShape("chevron", {
    x: 1.25,
    y: 4.75,
    w: 10.75,
    h: 0.7,
    line: { color: palette.blue, transparency: 100 },
    fill: { color: palette.blue, transparency: 26 },
  });

  [
    {
      x: 1.15,
      fillColor: palette.blue,
      title: "Mehr Lebenskompetenzen",
      body: "finanzielle Grundbildung, Medienkompetenz und alltagsrelevante Fähigkeiten stärker integrieren",
    },
    {
      x: 4.93,
      fillColor: palette.yellow,
      title: "Spätere Leistungsdifferenzierung",
      body: "mehr Zeit für Entwicklung und fundiertere Bildungsentscheidungen geben",
    },
    {
      x: 8.71,
      fillColor: palette.pink,
      title: "Alternative Bewertungsformen",
      body: "Lernfeedback, Portfolioarbeit und projektbasierte Bewertung ergänzen klassische Noten",
    },
  ].forEach((circle) => {
    slide.addShape("ellipse", {
      x: circle.x,
      y: 1.9,
      w: 2.6,
      h: 2.6,
      line: { color: palette.line, pt: 1 },
      fill: { color: circle.fillColor },
    });
    slide.addText(circle.title, {
      x: circle.x + 0.28,
      y: 2.45,
      w: 2.04,
      h: 0.7,
      fontFace: "Aptos Display",
      fontSize: 17,
      bold: true,
      color: palette.ink,
      align: "center",
      margin: 0,
      fit: "shrink",
    });
    slide.addText(circle.body, {
      x: circle.x + 0.18,
      y: 4.75,
      w: 2.25,
      h: 1.0,
      fontFace: "Aptos",
      fontSize: 13,
      color: palette.muted,
      align: "center",
      margin: 0,
      fit: "shrink",
    });
  });
  slide.addNotes(slideDefinitions[9].notes);
  registerSlideMeta(10, slideDefinitions[9].transition, slideDefinitions[9].animationTargets);
}

function buildSlide11(pptx) {
  const slide = pptx.addSlide();
  addSoftBackground(slide, 1);
  addEyebrow(slide, "Fazit");
  slide.addText("Schule beeinflusst mehr als nur Noten.", {
    x: 0.7,
    y: 1.1,
    w: 6.4,
    h: 0.95,
    fontFace: "Aptos Display",
    fontSize: 27,
    bold: true,
    color: palette.ink,
    margin: 0,
    fit: "shrink",
  });

  slide.addText(bulletRuns([
    "Strukturen wirken auf das Wohlbefinden.",
    "Leistungsdruck und Schulklima spielen eine zentrale Rolle.",
    "Reformen könnten Bildung und Glück stärker miteinander verbinden.",
  ], 20), {
    x: 0.95,
    y: 2.35,
    w: 5.8,
    h: 2.1,
    margin: 0,
  });
  slide.addText("Danke für eure Aufmerksamkeit!", {
    x: 0.95,
    y: 5.2,
    w: 4.9,
    h: 0.42,
    fontFace: "Aptos",
    fontSize: 18,
    color: palette.ink,
    bold: true,
    margin: 0,
  });
  addImagePanel(
    slide,
    path.join(ASSET_DIR, "calm-future-school.svg"),
    7.05,
    1.45,
    5.35,
    4.75,
    palette.yellow,
  );
  slide.addNotes(slideDefinitions[10].notes);
  registerSlideMeta(11, slideDefinitions[10].transition, slideDefinitions[10].animationTargets);
}

async function buildDeck() {
  ensureDirectories();
  writeAssets();
  writeNotesMarkdown();
  writeMetadata();

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Cursor";
  pptx.company = "Cursor";
  pptx.subject = "Präsentation zur Facharbeit im Seminarfachkurs Glück";
  pptx.title = "Wie Schule wirkt";
  pptx.lang = "de-DE";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "de-DE",
  };
  pptx.defineSlideMaster({
    title: "MASTER_PASTEL",
    background: { color: palette.white },
    objects: [],
    slideNumber: { x: 12.2, y: 6.85, fontFace: "Aptos", fontSize: 9, color: "9AA3AF" },
  });

  buildSlide1(pptx);
  buildSlide2(pptx);
  buildSlide3(pptx);
  buildSlide4(pptx);
  buildSlide5(pptx);
  buildSlide6(pptx);
  buildSlide7(pptx);
  buildSlide8(pptx);
  buildSlide9(pptx);
  buildSlide10(pptx);
  buildSlide11(pptx);

  const outFile = path.join(OUTPUT_DIR, "Wie_Schule_wirkt_Keynote_Style.pptx");
  await pptx.writeFile({ fileName: outFile });
  console.log(`Created ${outFile}`);
}

buildDeck().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
