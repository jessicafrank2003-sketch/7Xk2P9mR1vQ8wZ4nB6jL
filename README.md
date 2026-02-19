Prototyp: Automatisierte Stichprobenziehung (VBA)
Dieses Repository enthält den softwaregestützten Prototyp, der im Rahmen der Bachelorthesis „Konzeption und Implementierung einer automatisierten Lösung zur Optimierung der Stichprobenziehung im AFC-Kontext“ entwickelt wurde.

Projektbeschreibung
Der Prototyp dient als Proof-of-Concept für ein Schichtenmodell zur teilautomatisierten Selektion von Stichproben. Ziel ist die Steigerung der Prozesseffizienz und die Eliminierung manueller Fehlerquellen bei der ABC AG.

Technische Struktur
Die Anwendung ist modular nach einem Layer-Prinzip aufgebaut, um eine strikte Trennung von Zuständigkeiten zu gewährleisten:

UserInterface (UserForm): Die zentrale Interaktionsmaske für den Endanwender. Sie dient der Parametersteuerung (z. B. Zeitraumbegrenzung) und Fehlerabwicklung.

Hauptmodul (mod_Main): Das Orchestrierungs-Modul. Es steuert den Programmfluss und ruft die einzelnen Layer in der korrekten logischen Reihenfolge auf.

Layer 1 (Data): Verantwortlich für den Zugriff auf die Rohdaten und die Vorbereitung der Datenstrukturen.

Layer 2 (Logic): Enthält den Kern-Algorithmus zur statistischen oder risikobasierten Stichprobenziehung.

Layer 3 (Output/Reporting): Übernimmt die Aufbereitung der Ergebnisse und die Erstellung des Audit-Trails (Revisionssicherheit).
