#!/usr/bin/env python
#-*- coding: utf-8 -*-

#index - 
#This application is free software; you can redistribute
#it and/or modify it under the terms of the GNU General Public License
#defined in the COPYING file

#2010 Charlie Barnes.

import sys
import os
import gtk
import gobject
import mimetypes
import xlrd

class indexActions():
    def __init__(self):

        self.iec_scores = {
1: { "score" : 1, "taxa" : ["Abdera biflexuosa","Abdera biflexuosa"] },
2: { "score" : 3, "taxa" : ["Abdera quadrifasciata","Abdera quadrifasciata"] },
3: { "score" : 3, "taxa" : ["Abraeus granulum","Abraeus granulum"] },
4: { "score" : 3, "taxa" : ["Vanonus brevicornis","Aderus brevicornis"] },
5: { "score" : 1, "taxa" : ["Euglenes oculatus","Aderus oculatus"] },
6: { "score" : 3, "taxa" : ["Aeletes atomarius","Aeletes atomarius"] },
7: { "score" : 2, "taxa" : ["Agrilus (Anambus) biguttatus","Agrilus pannonicus"] },
8: { "score" : 3, "taxa" : ["Ampedus cardinalis","Ampedus cardinalis"] },
9: { "score" : 3, "taxa" : ["Ampedus cinnabarinus","Ampedus cinnabarinus"] },
10: { "score" : 1, "taxa" : ["Ampedus elongantulus","Ampedus elongatulus"] },
11: { "score" : 3, "taxa" : ["Ampedus nigerrimus","Ampedus nigerrimus"] },
12: { "score" : 1, "taxa" : ["Ampedus pomorum","Ampedus pomorum"] },
13: { "score" : 3, "taxa" : ["Ampedus quercicola","Ampedus quercicola (= pomonae)"] },
14: { "score" : 3, "taxa" : ["Ampedus rufipennis","Ampedus rufipennis"] },
15: { "score" : 1, "taxa" : ["Anisoxya fuscula","Anisoxya fuscula"] },
16: { "score" : 3, "taxa" : ["Anitys rubens","Anitys rubens"] },
17: { "score" : 1, "taxa" : ["Aplocnemus impressus","Aplocnemus impressus (=pini)"] },
18: { "score" : 1, "taxa" : ["Aplocnemus nigricornis","Aplocnemus nigricornis"] },
19: { "score" : 3, "taxa" : ["Atomaria (Atomaria) lohsei","Atomaria lohsei"] },
20: { "score" : 3, "taxa" : ["Aulonothroscus brevicollis","Aulonothroscus brevicollis"] },
21: { "score" : 3, "taxa" : ["Batrisodes adnexus","Batrisodes adnexus (=buqueti)"] },
22: { "score" : 3, "taxa" : ["Batrisodes delaporti","Batrisodes delaporti"] },
23: { "score" : 2, "taxa" : ["Batrisodes venustus","Batrisodes venustus"] },
24: { "score" : 2, "taxa" : ["Bibloporus minutus","Bibloporus minutus"] },
25: { "score" : 1, "taxa" : ["Biphyllus lunatus","Biphyllus lunatus"] },
26: { "score" : 1, "taxa" : ["Bitoma crenata","Bitoma crenata"] },
27: { "score" : 3, "taxa" : ["Brachygonus ruficeps","Brachygonus (= Ampedus) ruficeps"] },
28: { "score" : 1, "taxa" : ["Calambus bipustulatus","Calambus (= Selatosomus) bipustulatus"] },
29: { "score" : 1, "taxa" : ["Calosoma inquisitor","Calosoma inquisitor"] },
30: { "score" : 1, "taxa" : ["Carpophilus sexpustulatus","Carpophilus sexpustulatus"] },
31: { "score" : 1, "taxa" : ["Cerylon fagi","Cerylon fagi"] },
32: { "score" : 2, "taxa" : ["Cicones variegatus","Cicones variegata"] },
33: { "score" : 2, "taxa" : ["Orthocis coluber","Cis coluber"] },
34: { "score" : 3, "taxa" : ["Colydium elongatum","Colydium elongatum"] },
35: { "score" : 1, "taxa" : ["Conopalpus testaceus","Conopalpus testaceus"] },
36: { "score" : 3, "taxa" : ["Corticaria alleni","Corticaria alleni"] },
37: { "score" : 3, "taxa" : ["Corticaria fagi","Corticaria fagi"] },
38: { "score" : 3, "taxa" : ["Corticaria longicollis","Corticaria longicollis"] },
39: { "score" : 2, "taxa" : ["Corticeus unicolor","Corticeus unicolor"] },
40: { "score" : 1, "taxa" : ["Cossonus parallelepipedus","Cossonus parallelepipedus"] },
41: { "score" : 3, "taxa" : ["Cryptocephalus querceti","Cryptocephalus querceti"] },
42: { "score" : 3, "taxa" : ["Cryptophagus micaceus","Cryptophagus micaceus"] },
43: { "score" : 1, "taxa" : ["Ctesias serra","Ctesias serra"] },
44: { "score" : 2, "taxa" : ["Synchita separanda","Dienerella separanda"] },
45: { "score" : 2, "taxa" : ["Diplocoelus fagi","Diplocoelus fagi"] },
46: { "score" : 2, "taxa" : ["Dorcatoma chrysomelina","Dorcatoma chrysomelina"] },
47: { "score" : 2, "taxa" : ["Dorcatoma dresdensis","Dorcatoma dresdensis"] },
48: { "score" : 1, "taxa" : ["Dorcatoma flavicornis","Dorcatoma flavicornis"] },
49: { "score" : 2, "taxa" : ["Dorcatoma substriata","Dorcatoma serra"] },
50: { "score" : 3, "taxa" : ["Dryophthorus corticalis","Dryophthorus corticalis"] },
51: { "score" : 3, "taxa" : ["Elater ferrugineus","Elater ferrugineus"] },
52: { "score" : 1, "taxa" : ["Eledona agricola","Eledona agricola"] },
53: { "score" : 2, "taxa" : ["Enicmus brevicornis","Enicmus brevicornis"] },
54: { "score" : 2, "taxa" : ["Enicmus rugosus","Enicmus rugosus"] },
55: { "score" : 1, "taxa" : ["Epuraea (Epuraea) angustula","Epuraea angustula"] },
56: { "score" : 3, "taxa" : ["Ernoporicus caucasicus","Ernoporus caucasicus"] },
57: { "score" : 1, "taxa" : ["Ernoporicus fagi","Ernoporus fagi"] },
58: { "score" : 3, "taxa" : ["Eucnemis capucina","Eucnemis capucina"] },
59: { "score" : 3, "taxa" : ["Euconnus (Napochus) pragensis","Euconnus pragensis"] },
60: { "score" : 3, "taxa" : ["Euplectus brunneus","Euplectus brunneus"] },
61: { "score" : 3, "taxa" : ["Euplectus nanus","Euplectus nanus"] },
62: { "score" : 3, "taxa" : ["Euplectus punctatus","Euplectus punctatus"] },
63: { "score" : 2, "taxa" : ["Euryusa optabilis","Euryusa optabilis"] },
64: { "score" : 2, "taxa" : ["Euryusa sinuata","Euryusa sinuata"] },
65: { "score" : 3, "taxa" : ["Eutheia formicetorum","Eutheia formicetorum"] },
66: { "score" : 3, "taxa" : ["Eutheia linearis","Eutheia linearis"] },
67: { "score" : 3, "taxa" : ["Gastrallus immarginatus","Gastrallus immarginatus"] },
68: { "score" : 2, "taxa" : ["Globicornis nigripes","Globicornis rufitarsis (=nigripes)"] },
69: { "score" : 3, "taxa" : ["Gnorimus variabilis","Gnorimus variabilis"] },
70: { "score" : 3, "taxa" : ["Grammoptera ustulata","Grammoptera ustulata"] },
71: { "score" : 1, "taxa" : ["Grammoptera abdominalis","Grammoptera variegata"] },
72: { "score" : 1, "taxa" : ["Hallomenus binotatus","Hallomenus binotatus"] },
73: { "score" : 1, "taxa" : ["Hylecoetus dermestoides","Hylecoetus dermestoides"] },
74: { "score" : 3, "taxa" : ["Hypebaeus flavipes","Hypebaeus flavipes"] },
75: { "score" : 2, "taxa" : ["Hypulus quercinus","Hypulus quercinus"] },
76: { "score" : 2, "taxa" : ["Ischnodes sanguinicollis","Ischnodes sanguinicollis"] },
77: { "score" : 1, "taxa" : ["Ischnomera caerulea","Ischnomera caerulea"] },
78: { "score" : 3, "taxa" : ["Ischnomera cinerascens","Ischnomera cinerascens"] },
79: { "score" : 1, "taxa" : ["Ischnomera cyanea","Ischnomera cyanea"] },
80: { "score" : 3, "taxa" : ["Ischnomera sanguinicollis","Ischnomera sanguinicollis"] },
81: { "score" : 1, "taxa" : ["Korynetes caeruleus","Korynetes caeruleus"] },
82: { "score" : 3, "taxa" : ["Lacon querceus","Lacon quercus"] },
83: { "score" : 3, "taxa" : ["Laemophloeus monilis","Laemophloeus monilis"] },
84: { "score" : 3, "taxa" : ["Latridius consimilis","Lathridius consimilis"] },
85: { "score" : 1, "taxa" : ["Leptura aurulenta","Leptura (= Strangalia) aurulenta"] },
86: { "score" : 1, "taxa" : ["Leptura quadrifasciata","Leptura (= Strangalia) quadrifasciata"] },
87: { "score" : 3, "taxa" : ["Limoniscus violaceus","Limoniscus violaceus"] },
88: { "score" : 1, "taxa" : ["Lyctus brunneus","Lyctus brunneus"] },
89: { "score" : 3, "taxa" : ["Lymexylon navale","Lymexylon navale"] },
90: { "score" : 2, "taxa" : ["Malthodes crassicornis","Malthodes crassicornis"] },
91: { "score" : 3, "taxa" : ["Megapenthes lugens","Megapenthes lugens"] },
92: { "score" : 3, "taxa" : ["Melandrya barbata","Melandrya barbata"] },
93: { "score" : 1, "taxa" : ["Melandrya caraboides","Melandrya caraboides"] },
94: { "score" : 1, "taxa" : ["Melasis buprestoides","Melasis buprestoides"] },
95: { "score" : 1, "taxa" : ["Rhopalomesites tardyi","Mesites tardii"] },
96: { "score" : 2, "taxa" : ["Mesosa nebulosa","Mesosa nebulosa"] },
97: { "score" : 3, "taxa" : ["Micridium halidaii","Micridium halidaii"] },
98: { "score" : 1, "taxa" : ["Microrhagus pygmaeus","Microrhagus (= Dirhagus) pygmaeus"] },
99: { "score" : 3, "taxa" : ["Microscydmus minimus","Microscydmus minimus"] },
100: { "score" : 1, "taxa" : ["Mordella holomelaena","Mordella holomelaena (= aculeata)"] },
101: { "score" : 1, "taxa" : ["Mordella leucaspis","Mordella leucaspis  (= aculeata)"] },
102: { "score" : 1, "taxa" : ["Mycetochara humeralis","Mycetochara humeralis"] },
103: { "score" : 1, "taxa" : ["Mycetophagus atomarius","Mycetophagus atomarius"] },
104: { "score" : 1, "taxa" : ["Mycetophagus piceus","Mycetophagus piceus"] },
105: { "score" : 2, "taxa" : ["Notolaemus unifasciatus","Notolaemus unifasciatus"] },
106: { "score" : 1, "taxa" : ["Opilo mollis","Opilio mollis"] },
107: { "score" : 1, "taxa" : ["Orchesia undulata","Orchesia undulata"] },
108: { "score" : 2, "taxa" : ["Oxylaemus cylindricus","Oxylaemus cylindricus"] },
109: { "score" : 2, "taxa" : ["Oxylaemus variolosus","Oxylaemus variolosus"] },
110: { "score" : 2, "taxa" : ["Pediacus depressus","Pediacus depressus"] },
111: { "score" : 1, "taxa" : ["Pediacus dermestoides","Pediacus dermestoides"] },
112: { "score" : 2, "taxa" : ["Pedostrangalia revestita","Pedostrangalia (=Leptura) revestita"] },
113: { "score" : 1, "taxa" : ["Pentarthrum huttoni","Pentarthum huttoni"] },
114: { "score" : 1, "taxa" : ["Phloiophilus edwardsii","Phloiophilus edwardsi"] },
115: { "score" : 2, "taxa" : ["Phloiotrya vaudoueri","Phloiotrya vaudoueri"] },
116: { "score" : 3, "taxa" : ["Phyllodrepa nigra","Phyllodrepa nigra"] },
117: { "score" : 1, "taxa" : ["Phymatodes testaceus","Phymatodes testaceus"] },
118: { "score" : 3, "taxa" : ["Platycis cosnardi","Platycis cosnardi"] },
119: { "score" : 1, "taxa" : ["Platycis minutus","Platycis minutus"] },
120: { "score" : 1, "taxa" : ["Platypus cylindrus","Platypus cylindrus"] },
121: { "score" : 1, "taxa" : ["Platyrhinus resinosus","Platyrhinus resinosus"] },
122: { "score" : 1, "taxa" : ["Platystomos albinus","Platystomos albinus"] },
123: { "score" : 3, "taxa" : ["Plectophloeus nitidus","Plectophloeus nitidus"] },
124: { "score" : 2, "taxa" : ["Plegaderus dissectus","Plegaderus dissectus"] },
125: { "score" : 2, "taxa" : ["Prionocyphon serricornis","Prionocyphon serricornis"] },
126: { "score" : 1, "taxa" : ["Prionus coriarius","Prionus coriarius"] },
127: { "score" : 1, "taxa" : ["Prionychus ater","Prionychus ater"] },
128: { "score" : 3, "taxa" : ["Prionychus melanarius","Prionychus melanarius"] },
129: { "score" : 3, "taxa" : ["Procraerus tibialis","Procraerus tibialis"] },
130: { "score" : 2, "taxa" : ["Pseudocistela ceramboides","Pseudocistela ceramboides"] },
131: { "score" : 1, "taxa" : ["Pseudotriphyllus suturalis","Pseudotriphyllus suturalis"] },
132: { "score" : 2, "taxa" : ["Ptenidium (Gressnerium) gressneri","Ptenidium gressneri"] },
133: { "score" : 2, "taxa" : ["Ptenidium (Matthewsium) turgidum","Ptenidium turgidum"] },
134: { "score" : 2, "taxa" : ["Ptinella limbata","Ptinella limbata"] },
135: { "score" : 1, "taxa" : ["Ptinus palliatus","Ptinus palliatus"] },
136: { "score" : 2, "taxa" : ["Ptinus subpilosus","Ptinus subpilosus"] },
137: { "score" : 1, "taxa" : ["Pyrochroa coccinea","Pyrochroa coccinea"] },
138: { "score" : 2, "taxa" : ["Pyropterus nigroruber","Pyropterus nigroruber"] },
139: { "score" : 3, "taxa" : ["Pyrrhidium sanguineum","Pyrrhidium sanguineum"] },
140: { "score" : 1, "taxa" : ["Quedius (Microsaurus) aetolicus","Quedius aetolicus"] },
141: { "score" : 1, "taxa" : ["Quedius (Microsaurus) maurus","Quedius maurus"] },
142: { "score" : 1, "taxa" : ["Quedius (Microsaurus) microps","Quedius microps"] },
143: { "score" : 1, "taxa" : ["Quedius (Microsaurus) scitus","Quedius scitus"] },
144: { "score" : 1, "taxa" : ["Quedius (Microsaurus) truncicola","Quedius truncicola (=ventralis)"] },
145: { "score" : 1, "taxa" : ["Quedius (Microsaurus) xanthopus","Quedius xanthopus"] },
146: { "score" : 1, "taxa" : ["Rhizophagus (Rhizophagus) nitidulus","Rhizophagus nitidulus"] },
147: { "score" : 3, "taxa" : ["Rhizophagus (Rhizophagus) oblongicollis","Rhizophagus oblongicollis"] },
148: { "score" : 1, "taxa" : ["Saperda scalaris","Saperda scalaris"] },
149: { "score" : 3, "taxa" : ["Scraptia dubia","Scraptia dubia"] },
150: { "score" : 3, "taxa" : ["Scraptia fuscula","Scraptia fuscula"] },
151: { "score" : 3, "taxa" : ["Scraptia testacea","Scraptia testacea"] },
152: { "score" : 1, "taxa" : ["Scydmaenus (Cholerus) rufus","Scydmaenus rufus"] },
153: { "score" : 2, "taxa" : ["Silvanus bidentatus","Silvanus bidentatus"] },
154: { "score" : 1, "taxa" : ["Silvanus unidentatus","Silvanus unidentatus"] },
155: { "score" : 1, "taxa" : ["Sinodendron cylindricum","Sinodendron cylindricum"] },
156: { "score" : 1, "taxa" : ["Stenagostus rhombeus","Stenagostus rhombeus (= villosus)"] },
157: { "score" : 1, "taxa" : ["Stenichnus bicolor","Stenichnus bicolor"] },
158: { "score" : 3, "taxa" : ["Stenichnus godarti","Stenichnus godarti"] },
159: { "score" : 3, "taxa" : ["Stereocorynes truncorum","Stereocorynes (= Rhyncholus) truncorum"] },
160: { "score" : 3, "taxa" : ["Stictoleptura scutellata","Stictoleptura (=Anoplodera) scutellata"] },
161: { "score" : 1, "taxa" : ["Symbiotes latus","Symbiotes latus"] },
162: { "score" : 1, "taxa" : ["Synchita humeralis","Synchita humeralis"] },
163: { "score" : 3, "taxa" : ["Synchita separanda","Synchita separanda"] },
164: { "score" : 3, "taxa" : ["Tachyusida gracilis","Tachyusida gracilis"] },
165: { "score" : 3, "taxa" : ["Teredus cylindricus","Teredus cylindricus"] },
166: { "score" : 1, "taxa" : ["Tetratoma ancora","Tetratoma ancora"] },
167: { "score" : 1, "taxa" : ["Tetratoma desmarestii","Tetratoma desmaresti"] },
168: { "score" : 1, "taxa" : ["Tetratoma fungorum","Tetratoma fungorum"] },
169: { "score" : 1, "taxa" : ["Thanasimus formicarius","Thanasimus formicarius"] },
170: { "score" : 1, "taxa" : ["Thymalus limbatus","Thymalus limbatus"] },
171: { "score" : 1, "taxa" : ["Tillus elongatus","Tillus elongatus"] },
172: { "score" : 3, "taxa" : ["Tomoxia bucephala","Tomoxia bucephala (= biguttata)"] },
173: { "score" : 1, "taxa" : ["Trachodes hispidus","Trachodes hispidus"] },
174: { "score" : 2, "taxa" : ["Trichonyx sulcicollis","Trichonyx sulcicollis"] },
175: { "score" : 3, "taxa" : ["Trinodes hirtus","Trinodes hirtus"] },
176: { "score" : 1, "taxa" : ["Triphyllus bicolor","Triphyllus bicolor"] },
177: { "score" : 1, "taxa" : ["Triplax aenea","Triplax aenea"] },
178: { "score" : 1, "taxa" : ["Triplax lacordairii","Triplax lacordairii"] },
179: { "score" : 1, "taxa" : ["Triplax russica","Triplax russica"] },
180: { "score" : 1, "taxa" : ["Triplax scutellaris","Triplax scutellaris"] },
181: { "score" : 1, "taxa" : ["Tritoma bipustulata","Tritoma bipustulata"] },
182: { "score" : 1, "taxa" : ["Dissoleucas niveirostris","Tropideres niveirostris"] },
183: { "score" : 3, "taxa" : ["Enedreytes sepicola","Tropideres sepicola"] },
184: { "score" : 1, "taxa" : ["Trypodendron domesticum","Trypodendron (= Xyloterus) domesticum"] },
185: { "score" : 1, "taxa" : ["Trypodendron lineatum","Trypodendron (= Xyloterus) lineatum"] },
186: { "score" : 1, "taxa" : ["Trypodendron signatum","Trypodendron (= Xyloterus) signatum"] },
187: { "score" : 3, "taxa" : ["Uleiota planata","Uleiota planata"] },
188: { "score" : 1, "taxa" : ["Variimorda villosa","Variimorda villosa"] },
189: { "score" : 3, "taxa" : ["Velleius dilatatus","Velleius dilatatus"] },
190: { "score" : 1, "taxa" : ["Hypnogyra angularis","Xantholinus angularis"] },
191: { "score" : 1, "taxa" : ["Xestobium rufovillosum","Xestobium rufovillosum"] },
192: { "score" : 1, "taxa" : ["Xyleborinus saxesenii","Xyleborinus saxeseni"] },
193: { "score" : 1, "taxa" : ["Xyleborus dispar","Xyleborus dispar"] },
194: { "score" : 1, "taxa" : ["Xyleborus dryographus","Xyleborus dryographus"] },
195: { "score" : 1, "taxa" : ["Xyletinus longitarsis","Xyletinus longitarsus"] },
                          }

        self.riec_scores = { 
1: { "score" : 1, "taxa" : ["Abdera biflexuosa","Abdera biflexuosa"] },
2: { "score" : 3, "taxa" : ["Abdera quadrifasciata","Abdera quadrifasciata"] },
3: { "score" : 3, "taxa" : ["Abraeus granulum","Abraeus granulum"] },
4: { "score" : 3, "taxa" : ["Vanonus brevicornis","Aderus brevicornis"] },
5: { "score" : 1, "taxa" : ["Euglenes oculatus","Aderus oculatus"] },
6: { "score" : 3, "taxa" : ["Aeletes atomarius","Aeletes atomarius"] },
7: { "score" : 3, "taxa" : ["Ampedus cardinalis","Ampedus cardinalis"] },
8: { "score" : 3, "taxa" : ["Ampedus cinnabarinus","Ampedus cinnabarinus"] },
9: { "score" : 1, "taxa" : ["Ampedus elongantulus","Ampedus elongatulus"] },
10: { "score" : 3, "taxa" : ["Ampedus nigerrimus","Ampedus nigerrimus"] },
11: { "score" : 1, "taxa" : ["Ampedus pomorum","Ampedus pomorum"] },
12: { "score" : 3, "taxa" : ["Ampedus quercicola","Ampedus quercicola (= pomonae)"] },
13: { "score" : 3, "taxa" : ["Ampedus rufipennis","Ampedus rufipennis"] },
14: { "score" : 3, "taxa" : ["Anaspis (Anaspis) thoracica","Anaspis septentrionalis (= schilskyana)"] },
15: { "score" : 1, "taxa" : ["Anisoxya fuscula","Anisoxya fuscula"] },
16: { "score" : 3, "taxa" : ["Anitys rubens","Anitys rubens"] },
17: { "score" : 2, "taxa" : ["Anoplodera sexguttata","Anoplodera (= Leptura) sexguttata"] },
18: { "score" : 2, "taxa" : ["Aplocnemus impressus","Aplocnemus impressus (=pini)"] },
19: { "score" : 2, "taxa" : ["Aplocnemus nigricornis","Aplocnemus nigricornis"] },
20: { "score" : 3, "taxa" : ["Aulonothroscus brevicollis","Aulonothroscus brevicollis"] },
21: { "score" : 3, "taxa" : ["Batrisodes adnexus","Batrisodes adnexus (=buqueti)"] },
22: { "score" : 3, "taxa" : ["Batrisodes delaporti","Batrisodes delaporti"] },
23: { "score" : 3, "taxa" : ["Batrisodes venustus","Batrisodes venustus"] },
24: { "score" : 2, "taxa" : ["Bibloporus minutus","Bibloporus minutus"] },
25: { "score" : 1, "taxa" : ["Biphyllus lunatus","Biphyllus lunatus"] },
26: { "score" : 1, "taxa" : ["Bitoma crenata","Bitoma crenata"] },
27: { "score" : 3, "taxa" : ["Brachygonus ruficeps","Brachygonus (= Ampedus) ruficeps"] },
28: { "score" : 1, "taxa" : ["Calambus bipustulatus","Calambus (= Selatosomus) bipustulatus"] },
29: { "score" : 1, "taxa" : ["Carpophilus sexpustulatus","Carpophilus sexpustulatus"] },
30: { "score" : 2, "taxa" : ["Cerylon fagi","Cerylon fagi"] },
31: { "score" : 2, "taxa" : ["Cicones variegatus","Cicones variegata"] },
32: { "score" : 2, "taxa" : ["Orthocis coluber","Cis coluber"] },
33: { "score" : 1, "taxa" : ["Conopalpus testaceus","Conopalpus testaceus"] },
34: { "score" : 3, "taxa" : ["Corticaria alleni","Corticaria alleni"] },
35: { "score" : 2, "taxa" : ["Corticeus unicolor","Corticeus unicolor"] },
36: { "score" : 1, "taxa" : ["Cossonus parallelepipedus","Cossonus parallelepipedus"] },
37: { "score" : 3, "taxa" : ["Cryptophagus micaceus","Cryptophagus micaceus"] },
38: { "score" : 1, "taxa" : ["Diplocoelus fagi","Diplocoelus fagi"] },
39: { "score" : 2, "taxa" : ["Dorcatoma ambjoerni","Dorcatoma ambjourni"] },
40: { "score" : 1, "taxa" : ["Dorcatoma chrysomelina","Dorcatoma chrysomelina"] },
41: { "score" : 2, "taxa" : ["Dorcatoma dresdensis","Dorcatoma dresdensis"] },
42: { "score" : 1, "taxa" : ["Dorcatoma flavicornis","Dorcatoma flavicornis"] },
43: { "score" : 2, "taxa" : ["Dorcatoma substriata","Dorcatoma serra"] },
44: { "score" : 3, "taxa" : ["Dryophthorus corticalis","Dryophthorus corticalis"] },
45: { "score" : 3, "taxa" : ["Elater ferrugineus","Elater ferrugineus"] },
46: { "score" : 1, "taxa" : ["Eledona agricola","Eledona agricola"] },
47: { "score" : 1, "taxa" : ["Enicmus brevicornis","Enicmus brevicornis"] },
48: { "score" : 2, "taxa" : ["Enicmus rugosus","Enicmus rugosus"] },
49: { "score" : 1, "taxa" : ["Epuraea (Epuraea) angustula","Epuraea angustula"] },
50: { "score" : 2, "taxa" : ["Ernoporicus caucasicus","Ernoporus caucasicus"] },
51: { "score" : 1, "taxa" : ["Ernoporicus fagi","Ernoporus fagi"] },
52: { "score" : 2, "taxa" : ["Ernoporus tiliae","Ernoporus tiliae"] },
53: { "score" : 3, "taxa" : ["Eucnemis capucina","Eucnemis capucina"] },
54: { "score" : 3, "taxa" : ["Euconnus (Napochus) pragensis","Euconnus pragensis"] },
55: { "score" : 3, "taxa" : ["Euplectus nanus","Euplectus nanus"] },
56: { "score" : 3, "taxa" : ["Euplectus punctatus","Euplectus punctatus"] },
57: { "score" : 2, "taxa" : ["Euryusa optabilis","Euryusa optabilis"] },
58: { "score" : 2, "taxa" : ["Euryusa sinuata","Euryusa sinuata"] },
59: { "score" : 3, "taxa" : ["Eutheia formicetorum","Eutheia formicetorum"] },
60: { "score" : 3, "taxa" : ["Eutheia linearis","Eutheia linearis"] },
61: { "score" : 3, "taxa" : ["Gastrallus immarginatus","Gastrallus immarginatus"] },
62: { "score" : 3, "taxa" : ["Globicornis nigripes","Globicornis rufitarsis (=nigripes)"] },
63: { "score" : 3, "taxa" : ["Gnorimus nobilis","Gnorimus nobilis"] },
64: { "score" : 3, "taxa" : ["Gnorimus variabilis","Gnorimus variabilis"] },
65: { "score" : 3, "taxa" : ["Grammoptera ustulata","Grammoptera ustulata"] },
66: { "score" : 1, "taxa" : ["Grammoptera abdominalis","Grammoptera variegata"] },
67: { "score" : 1, "taxa" : ["Hallomenus binotatus","Hallomenus binotatus"] },
68: { "score" : 1, "taxa" : ["Hylecoetus dermestoides","Hylecoetus dermestoides"] },
69: { "score" : 3, "taxa" : ["Hypebaeus flavipes","Hypebaeus flavipes"] },
70: { "score" : 3, "taxa" : ["Hypulus quercinus","Hypulus quercinus"] },
71: { "score" : 2, "taxa" : ["Ischnodes sanguinicollis","Ischnodes sanguinicollis"] },
72: { "score" : 3, "taxa" : ["Ischnomera caerulea","Ischnomera caerulea"] },
73: { "score" : 1, "taxa" : ["Ischnomera cinerascens","Ischnomera cinerascens"] },
74: { "score" : 1, "taxa" : ["Ischnomera cyanea","Ischnomera cyanea"] },
75: { "score" : 3, "taxa" : ["Ischnomera sanguinicollis","Ischnomera sanguinicollis"] },
76: { "score" : 1, "taxa" : ["Korynetes caeruleus","Korynetes caeruleus"] },
77: { "score" : 3, "taxa" : ["Lacon querceus","Lacon quercus"] },
78: { "score" : 3, "taxa" : ["Latridius consimilis","Lathridius consimilis"] },
79: { "score" : 1, "taxa" : ["Leptura aurulenta","Leptura (= Strangalia) aurulenta"] },
80: { "score" : 1, "taxa" : ["Leptura quadrifasciata","Leptura (= Strangalia) quadrifasciata"] },
81: { "score" : 3, "taxa" : ["Limoniscus violaceus","Limoniscus violaceus"] },
82: { "score" : 1, "taxa" : ["Lyctus brunneus","Lyctus brunneus"] },
83: { "score" : 2, "taxa" : ["Lymexylon navale","Lymexylon navale"] },
84: { "score" : 3, "taxa" : ["Malthodes crassicornis","Malthodes crassicornis"] },
85: { "score" : 3, "taxa" : ["Megapenthes lugens","Megapenthes lugens"] },
86: { "score" : 3, "taxa" : ["Melandrya barbata","Melandrya barbata"] },
87: { "score" : 1, "taxa" : ["Melandrya caraboides","Melandrya caraboides"] },
88: { "score" : 1, "taxa" : ["Melasis buprestoides","Melasis buprestoides"] },
89: { "score" : 1, "taxa" : ["Rhopalomesites tardyi","Mesites tardii"] },
90: { "score" : 2, "taxa" : ["Mesosa nebulosa","Mesosa nebulosa"] },
91: { "score" : 3, "taxa" : ["Micridium halidaii","Micridium halidaii"] },
92: { "score" : 1, "taxa" : ["Microrhagus pygmaeus","Microrhagus (= Dirhagus) pygmaeus"] },
93: { "score" : 3, "taxa" : ["Microscydmus minimus","Microscydmus minimus"] },
94: { "score" : 2, "taxa" : ["Microscydmus nanus","Microscydmus nanus"] },
95: { "score" : 1, "taxa" : ["Mordellistena (Mordellistena) neuwaldeggiana","Mordellistena neuwaldeggiana"] },
96: { "score" : 2, "taxa" : ["Mycetochara humeralis","Mycetochara humeralis"] },
97: { "score" : 1, "taxa" : ["Mycetophagus atomarius","Mycetophagus atomarius"] },
98: { "score" : 2, "taxa" : ["Mycetophagus piceus","Mycetophagus piceus"] },
99: { "score" : 2, "taxa" : ["Mycetophagus populi","Mycetophagus populi"] },
100: { "score" : 2, "taxa" : ["Mycetophagus quadriguttatus","Mycetophagus quadriguttatus"] },
101: { "score" : 2, "taxa" : ["Notolaemus unifasciatus","Notolaemus unifasciatus"] },
102: { "score" : 1, "taxa" : ["Opilo mollis","Opilio mollis"] },
103: { "score" : 1, "taxa" : ["Orchesia undulata","Orchesia undulata"] },
104: { "score" : 2, "taxa" : ["Oxylaemus variolosus","Oxylaemus variolosus"] },
105: { "score" : 2, "taxa" : ["Pediacus depressus","Pediacus depressus"] },
106: { "score" : 1, "taxa" : ["Pediacus dermestoides","Pediacus dermestoides"] },
107: { "score" : 2, "taxa" : ["Pedostrangalia revestita","Pedostrangalia (=Leptura) revestita"] },
108: { "score" : 1, "taxa" : ["Phloiophilus edwardsii","Phloiophilus edwardsi"] },
109: { "score" : 2, "taxa" : ["Phloiotrya vaudoueri","Phloiotrya vaudoueri"] },
110: { "score" : 3, "taxa" : ["Phyllodrepa nigra","Phyllodrepa nigra"] },
111: { "score" : 1, "taxa" : ["Phymatodes testaceus","Phymatodes testaceus"] },
112: { "score" : 3, "taxa" : ["Platycis cosnardi","Platycis cosnardi"] },
113: { "score" : 1, "taxa" : ["Platycis minutus","Platycis minutus"] },
114: { "score" : 1, "taxa" : ["Platypus cylindrus","Platypus cylindrus"] },
115: { "score" : 1, "taxa" : ["Platyrhinus resinosus","Platyrhinus resinosus"] },
116: { "score" : 1, "taxa" : ["Platystomos albinus","Platystomos albinus"] },
117: { "score" : 3, "taxa" : ["Plectophloeus nitidus","Plectophloeus nitidus"] },
118: { "score" : 2, "taxa" : ["Plegaderus dissectus","Plegaderus dissectus"] },
119: { "score" : 1, "taxa" : ["Prionocyphon serricornis","Prionocyphon serricornis"] },
120: { "score" : 1, "taxa" : ["Prionus coriarius","Prionus coriarius"] },
121: { "score" : 1, "taxa" : ["Prionychus ater","Prionychus ater"] },
122: { "score" : 3, "taxa" : ["Prionychus melanarius","Prionychus melanarius"] },
123: { "score" : 3, "taxa" : ["Procraerus tibialis","Procraerus tibialis"] },
124: { "score" : 2, "taxa" : ["Pseudocistela ceramboides","Pseudocistela ceramboides"] },
125: { "score" : 1, "taxa" : ["Pseudotriphyllus suturalis","Pseudotriphyllus suturalis"] },
126: { "score" : 2, "taxa" : ["Ptenidium (Gressnerium) gressneri","Ptenidium gressneri"] },
127: { "score" : 2, "taxa" : ["Ptenidium (Matthewsium) turgidum","Ptenidium turgidum"] },
128: { "score" : 2, "taxa" : ["Ptinella limbata","Ptinella limbata"] },
129: { "score" : 2, "taxa" : ["Ptinus subpilosus","Ptinus subpilosus"] },
130: { "score" : 1, "taxa" : ["Pyrochroa coccinea","Pyrochroa coccinea"] },
131: { "score" : 1, "taxa" : ["Pyropterus nigroruber","Pyropterus nigroruber"] },
132: { "score" : 3, "taxa" : ["Pyrrhidium sanguineum","Pyrrhidium sanguineum"] },
133: { "score" : 1, "taxa" : ["Quedius (Microsaurus) aetolicus","Quedius aetolicus"] },
134: { "score" : 1, "taxa" : ["Quedius (Microsaurus) maurus","Quedius maurus"] },
135: { "score" : 1, "taxa" : ["Quedius (Microsaurus) microps","Quedius microps"] },
136: { "score" : 2, "taxa" : ["Quedius (Microsaurus) scitus","Quedius scitus"] },
137: { "score" : 1, "taxa" : ["Quedius (Microsaurus) truncicola","Quedius truncicola (=ventralis)"] },
138: { "score" : 1, "taxa" : ["Quedius (Microsaurus) xanthopus","Quedius xanthopus"] },
139: { "score" : 1, "taxa" : ["Rhizophagus (Rhizophagus) nitidulus","Rhizophagus nitidulus"] },
140: { "score" : 3, "taxa" : ["Rhizophagus (Rhizophagus) oblongicollis","Rhizophagus oblongicollis"] },
141: { "score" : 1, "taxa" : ["Saperda scalaris","Saperda scalaris"] },
142: { "score" : 3, "taxa" : ["Scraptia fuscula","Scraptia fuscula"] },
143: { "score" : 3, "taxa" : ["Scraptia testacea","Scraptia testacea"] },
144: { "score" : 1, "taxa" : ["Scydmaenus (Cholerus) rufus","Scydmaenus rufus"] },
145: { "score" : 2, "taxa" : ["Silvanus bidentatus","Silvanus bidentatus"] },
146: { "score" : 1, "taxa" : ["Silvanus unidentatus","Silvanus unidentatus"] },
147: { "score" : 1, "taxa" : ["Stenagostus rhombeus","Stenagostus rhombeus (= villosus)"] },
148: { "score" : 1, "taxa" : ["Stenichnus bicolor","Stenichnus bicolor"] },
149: { "score" : 2, "taxa" : ["Stenichnus godarti","Stenichnus godarti"] },
150: { "score" : 3, "taxa" : ["Stereocorynes truncorum","Stereocorynes (= Rhyncholus) truncorum"] },
151: { "score" : 3, "taxa" : ["Stictoleptura scutellata","Stictoleptura (=Anoplodera) scutellata"] },
152: { "score" : 1, "taxa" : ["Symbiotes latus","Symbiotes latus"] },
153: { "score" : 1, "taxa" : ["Synchita humeralis","Synchita humeralis"] },
154: { "score" : 1, "taxa" : ["Synchita separanda","Synchita separanda"] },
155: { "score" : 3, "taxa" : ["Tachyusida gracilis","Tachyusida gracilis"] },
156: { "score" : 3, "taxa" : ["Teredus cylindricus","Teredus cylindricus"] },
157: { "score" : 1, "taxa" : ["Tetratoma ancora","Tetratoma ancora"] },
158: { "score" : 1, "taxa" : ["Tetratoma desmarestii","Tetratoma desmaresti"] },
159: { "score" : 1, "taxa" : ["Thanasimus formicarius","Thanasimus formicarius"] },
160: { "score" : 2, "taxa" : ["Thymalus limbatus","Thymalus limbatus"] },
161: { "score" : 1, "taxa" : ["Tillus elongatus","Tillus elongatus"] },
162: { "score" : 1, "taxa" : ["Tomoxia bucephala","Tomoxia bucephala (= biguttata)"] },
163: { "score" : 1, "taxa" : ["Trachodes hispidus","Trachodes hispidus"] },
164: { "score" : 3, "taxa" : ["Trinodes hirtus","Trinodes hirtus"] },
165: { "score" : 2, "taxa" : ["Triphyllus bicolor","Triphyllus bicolor"] },
166: { "score" : 1, "taxa" : ["Triplax lacordairii","Triplax lacordairii"] },
167: { "score" : 1, "taxa" : ["Triplax russica","Triplax russica"] },
168: { "score" : 1, "taxa" : ["Triplax scutellaris","Triplax scutellaris"] },
169: { "score" : 1, "taxa" : ["Tritoma bipustulata","Tritoma bipustulata"] },
170: { "score" : 1, "taxa" : ["Dissoleucas niveirostris","Tropideres niveirostris"] },
171: { "score" : 3, "taxa" : ["Enedreytes sepicola","Tropideres sepicola"] },
172: { "score" : 1, "taxa" : ["Trypodendron domesticum","Trypodendron (= Xyloterus) domesticum"] },
173: { "score" : 1, "taxa" : ["Trypodendron signatum","Trypodendron (= Xyloterus) signatum"] },
174: { "score" : 2, "taxa" : ["Uleiota planata","Uleiota planata"] },
175: { "score" : 3, "taxa" : ["Velleius dilatatus","Velleius dilatatus"] },
176: { "score" : 2, "taxa" : ["Hypnogyra angularis","Xantholinus angularis"] },
177: { "score" : 1, "taxa" : ["Xestobium rufovillosum","Xestobium rufovillosum"] },
178: { "score" : 1, "taxa" : ["Xyleborinus saxesenii","Xyleborinus saxeseni"] },
179: { "score" : 1, "taxa" : ["Xyleborus dispar","Xyleborus dispar"] },
180: { "score" : 1, "taxa" : ["Xyleborus dryographus","Xyleborus dryographus"] },
                           }
                        
        #Load the widget tree
        builder = ""
        self.builder = gtk.Builder()
        self.builder.add_from_string(builder, len(builder))
        self.builder.add_from_file("ui.xml")

        signals = {
                   "mainQuit":self.main_quit,
                   "showAboutDialog":self.show_about_dialog,
                   "parse":self.parse,
                   "selectFile":self.select_file,
                  }
        self.builder.connect_signals(signals)

        treeview = self.builder.get_object("treeview1")
        model = gtk.ListStore(str, int, int, int, int, int)
        treeview.set_headers_visible(True)
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Site", cell, text=0)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(0)
        treeview.append_column(column)
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Species", cell, text=1)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(1)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("IEC Scoring Species", cell, text=2)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(2)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("IEC", cell, text=3)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(3)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("RIEC Scoring Species", cell, text=4)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(4)
        treeview.append_column(column)    
        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn("RIEC", cell, text=5)
        column.set_resizable(True)
        column.set_expand(False)
        column.set_sort_column_id(5)
        treeview.append_column(column)    
            
        treeview.set_model(model)

        #Setup the main window
        self.main_window = self.builder.get_object("window1")
        self.main_window.show()
              
    def select_file(self, widget):
        filetype = mimetypes.guess_type(self.builder.get_object("filechooserbutton2").get_filename())[0]
        
        if filetype == "application/vnd.ms-excel":
            self.parse(widget)
              
    def parse(self, widget):

        cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
        self.builder.get_object("window1").window.set_cursor(cursor)
    
        while gtk.events_pending():
            gtk.main_iteration()
                    
        treeview = self.builder.get_object("treeview1")
        model = treeview.get_model()
        model.clear()
        
        filename = self.builder.get_object("filechooserbutton2").get_filename()
        filetype = mimetypes.guess_type(filename)[0]
        
        if filetype == "application/vnd.ms-excel":
            book = xlrd.open_workbook(filename)
            
            if book.nsheets > 1:
                
                dialog = self.builder.get_object("dialog1")

                try:
                    self.builder.get_object("hbox5").get_children()[1].destroy()           
                except IndexError:
                    pass
                    
                combobox = gtk.combo_box_new_text()
                
                for name in book.sheet_names():
                    combobox.append_text(name)
                    
                combobox.set_active(0)
                combobox.show()
                self.builder.get_object("hbox5").add(combobox)
                
                self.builder.get_object("window1").window.set_cursor(None)
            
                while gtk.events_pending():
                    gtk.main_iteration()
                
                response = dialog.run()

                if response == 1:
                    sheet = book.sheet_by_name(combobox.get_active_text())
                else:
                    dialog.hide()
                    return -1
                    
                dialog.hide()
                
            else:
                sheet = book.sheet_by_index(0)

            self.builder.get_object("vbox1").set_sensitive(False)
            
            cursor = gtk.gdk.Cursor(gtk.gdk.WATCH)
            self.builder.get_object("window1").window.set_cursor(cursor)
        
            while gtk.events_pending():
                gtk.main_iteration()
                
            for col_index in range(sheet.ncols):
                if sheet.cell(0, col_index).value == "Site":
                    site_position = col_index
                elif sheet.cell(0, col_index).value.lower() == "location":
                    site_position = col_index
                elif sheet.cell(0, col_index).value == "Species":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Taxon Name":
                    taxon_position = col_index
                elif sheet.cell(0, col_index).value == "Date":
                    date_position = col_index

            data = {}
            
            for row_index in range(1, sheet.nrows):
                site = sheet.cell(row_index, site_position).value
                taxon = sheet.cell(row_index, taxon_position).value
                    
                if data.has_key(site) and taxon not in data[site]["species_list"]:
                    data[site]["species_list"].append(taxon)
                elif not data.has_key(site):
                    data[site] = { }
                    data[site]["species_list"] = [taxon, ]
                    data[site]["iec_scoring_species"] = [ ]
                    data[site]["riec_scoring_species"] = [ ]
                    data[site]["iec_score"] = 0
                    data[site]["riec_score"] = 0
                    
            self.builder.get_object("progressbar1").show()

            count = 0.0
            total = len(data)

            for site in data:                    
                for taxon in data[site]["species_list"]:
                    for code in self.iec_scores:
                        if taxon in self.iec_scores[code]["taxa"]:
                            if code not in data[site]["iec_scoring_species"]:
                                data[site]["iec_scoring_species"].append(code)
                                data[site]["iec_score"] = data[site]["iec_score"] + self.iec_scores[code]["score"]
                                
                    for code in self.riec_scores:
                        if taxon in self.riec_scores[code]["taxa"]:
                            if code not in data[site]["riec_scoring_species"]:
                                data[site]["riec_scoring_species"].append(code)
                                data[site]["riec_score"] = data[site]["riec_score"] + self.riec_scores[code]["score"]
       
                model.append([site, len(data[site]["species_list"]), len(data[site]["iec_scoring_species"]), data[site]["iec_score"], len(data[site]["riec_scoring_species"]), data[site]["riec_score"]])

                self.builder.get_object("progressbar1").set_fraction(count/total)
                self.builder.get_object("progressbar1").set_text(''.join(["Processed ", str(int(count)), " of ", str(total), " sites"]))
                count = count + 1.0
                
                while gtk.events_pending():
                    gtk.main_iteration()

        self.builder.get_object("progressbar1").hide()
        self.builder.get_object("window1").window.set_cursor(None)
        self.builder.get_object("vbox1").set_sensitive(True)
        
        while gtk.events_pending():
            gtk.main_iteration()
                
    def main_quit(self, widget, var=None):
        gtk.main_quit()

    def show_about_dialog(self, widget):
       about=gtk.AboutDialog()
       about.set_name("indec")
       about.set_copyright("2010 Charlie Barnes")
       about.set_authors(["Charlie Barnes <charlie@cucaera.co.uk>"])
       about.set_license("indec is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nindec is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with indec; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
       about.set_wrap_license(True)
       about.set_website("http://cucaera.co.uk/software/indec/")
       about.set_transient_for(self.builder.get_object("window1"))
       result=about.run()
       about.destroy()

if __name__ == '__main__':
    indexActions()
    gtk.main()
    
