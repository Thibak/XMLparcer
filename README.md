XMLparcer
=========

Формат корпуса BNC

Корпус хранится в некоторой директории

BNCpath =

Структура директории корпуса неизменна, за исключением одного нехватающего файла

BNC_WORLD_INDEX.XLS

Который мы туда помещаем. О применении и синтаксисе ниже.

 - BNCpath
   - Doc
      - HTML (user guide)
      -Src (user guide)
   - Etc
      -file_index
 - Index
      - xid
      -xgrammar
 - install
 - Texts (A-K)
 - Usr
 - XML
      - scrips
  - BNC_WORLD_INDEX
  - bncHdr
  - bncBib
  - corpus_parameters
         
=======

нужные в таблице данные: слово, левое и правое окружение, класс, возраст, пол, роль, тип взаимодействия, источник (?)
отбор данных: только устный, исключая передачи
можно задавать в параметрах - устный, класс и существ. переменная - класс (AB, C1, C2, DE, uncl)


===

запрос: if Q="S" and C="S_Demog_AB" or "S_Demog_C1" or "S_Demog_C2" or "S_Demog_DE" or "S_Demog_Unclassified"

нужные столбцы: A, C, M, R, S, T

+ роль, адресат


===
часть речи: NN0, NN1, NN1-AJ0, NN1-NP0, NN1-WB, NN1-WG, NN2, NN2-WZ, NP0-NN1, UNC, WB-NN1, WG-NN1, WZ-NN2
