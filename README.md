# Delta Memory Force 2.0

**Delta Memory Force 2.0** este o aplicație Windows pentru monitorizarea și administrarea resurselor sistemului, realizată în perioada **1999-2001**. Ultima versiune a fost compilată în jurul anului **2004**. Proiectul este publicat în prezent cu scop istoric, educațional și arhivistic, ca exemplu de aplicație de sistem dezvoltată pentru generațiile Windows de la sfârșitul anilor 1990 și începutul anilor 2000.

Prima versiune distribuită public, Delta Memory Force 1.0, a apărut la începutul anilor 2000 pe compilația internațională <a href="https://archive.org/details/onyxdvd-11">Onyx DVD #11 – Adrenaline</a>, sub denumirea „Delta Memory Force 1.0 (c) RamForces”. Această includere reprezintă cea mai veche apariție publică documentată a aplicației identificată până în prezent.

Programul afișează în timp real memoria RAM liberă și totală, utilizarea procesorului, graficele de activitate pentru CPU și memorie, lista proceselor active, prioritatea și numărul de fire de execuție ale procesului selectat, precum și capacitatea și spațiul liber ale unităților de stocare.

Aplicația permite eliberarea memoriei RAM, schimbarea priorității proceselor, terminarea acestora și accesarea unor funcții avansate de administrare. Interfața grafică, specifică începutului anilor 2000, folosește un stil tehnic distinct, cu fundal negru, grafice verzi și indicatori colorați actualizați în timp real.

![Screenshot](https://github.com/Gagniuc/ATM-Lab/blob/main/img/DMF_open.jpg)

Delta Memory Force elibera memoria RAM prin alocarea temporară a unei cantități mari de memorie. Această operație determina Windows să mute în memoria virtuală, pe disc, paginile inactive ale proceselor și să reducă memoria fizică ocupată. După încheierea operației, memoria temporar alocată de program era eliberată, rezultând o creștere imediată a cantității de RAM disponibile.

Programul nu ștergea date și nu închidea aplicațiile, ci forța sistemul de operare să reorganizeze memoria, păstrând în RAM în special datele active și transferând pe disc informațiile utilizate mai rar. Acest principiu era cunoscut la începutul anilor 2000 sub denumirea de RAM optimization sau memory forcing și explică numele aplicației, Delta Memory Force.

Aplicația a fost concepută pentru generațiile Windows din acea perioadă, cel mai probabil Windows 98, Windows NT și Windows 2000, fiind posibil să funcționeze și pe Windows XP. Compatibilitatea cu sistemele Windows moderne nu este garantată.

## Funcționalități

Aplicația afișează în timp real:

- memoria RAM liberă și memoria totală;
- gradul de utilizare a procesorului;
- grafice pentru activitatea CPU și a memoriei;
- lista proceselor active;
- prioritatea procesului selectat;
- numărul de fire de execuție ale procesului selectat;
- capacitatea și spațiul liber ale unităților de stocare;
- numărul total de procese active.

Programul permite, de asemenea:

- eliberarea memoriei RAM;
- schimbarea priorității proceselor;
- terminarea proceselor selectate;
- actualizarea listei de procese;
- accesarea unor opțiuni suplimentare de administrare a sistemului.

## Principiul de eliberare a memoriei

Delta Memory Force elibera memoria RAM prin **alocarea temporară a unei cantități mari de memorie**.

Această operație determina sistemul Windows să reducă memoria fizică ocupată de paginile inactive și să transfere o parte dintre acestea în memoria virtuală, pe disc. După finalizarea procesului, memoria alocată temporar de aplicație era eliberată, rezultând o creștere imediată a cantității de RAM raportate ca fiind disponibile.

Programul nu ștergea date și nu închidea automat aplicațiile. El forța sistemul de operare să reorganizeze utilizarea memoriei, păstrând în RAM în special datele active și mutând temporar pe disc informațiile utilizate mai rar.

La începutul anilor 2000, această tehnică era cunoscută sub denumiri precum:

- **RAM optimization**;
- **memory forcing**;
- **memory freeing**.

Acest principiu explică și numele aplicației: **Delta Memory Force**.

> Creșterea valorii „Free RAM” nu înseamnă neapărat o creștere permanentă a performanței. În anumite situații, accesarea ulterioară a datelor mutate în memoria virtuală putea produce activitate suplimentară pe disc.

## Interfață

Interfața grafică reflectă stilul aplicațiilor tehnice de la începutul anilor 2000:

- fundal negru;
- grafice și texte verzi;
- indicatori colorați actualizați în timp real;
- controale compacte pentru monitorizarea rapidă a sistemului;
- panou extins pentru procese, prioritate și fire de execuție.

## Compatibilitate

Aplicația a fost proiectată pentru sistemele Windows disponibile în perioada în care a fost dezvoltată.

Compatibilitatea probabilă include:

- Windows 98;
- Windows NT;
- Windows 2000;
- posibil Windows XP.

Compatibilitatea exactă urmează să fie verificată prin analizarea executabilului original și prin testare în mașini virtuale.

Aplicația nu este destinată sistemelor Windows moderne, iar funcționarea sa pe Windows 10 sau Windows 11 nu este garantată.

## Cod sursă

Intenționez să public și **codul sursă original al aplicației**, în măsura în care acesta poate fi recuperat și organizat.

Codul va fi oferit ca:

- material istoric;
- exemplu de programare Windows;
- exemplu de monitorizare a resurselor sistemului;
- exemplu de administrare a proceselor;
- proiect software reprezentativ pentru începutul anilor 2000.

## Starea proiectului

- **Perioada inițială de dezvoltare:** 2000–2001
- **Ultima versiune compilată:** aproximativ 2004
- **Stare actuală:** proiect istoric / arhivistic
- **Platformă originală:** Microsoft Windows
- **Versiune prezentată:** Delta Memory Force 2.0

## Notă

Acest proiect este publicat pentru documentare și conservare istorică. Unele funcții de administrare a proceselor și memoriei pot necesita drepturi speciale sau pot avea un comportament diferit pe sisteme de operare mai noi.

Nu se recomandă executarea aplicației originale pe un sistem de producție. Pentru testare, este indicată utilizarea unei mașini virtuale izolate.
