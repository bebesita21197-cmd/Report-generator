v1.15.12 changes
- Main header redesign: animated gradient + floating glow orbs, shimmering title, spinning logo, glass "Created by" chip with avatar.

v1.15.11 changes
- Generate buttons par cursor le jane par button blank (white) ho jata tha - fix, ab hover par apna colour rehta hai.

v1.15.10 changes
- "Done / Report ready" line aur saare count blocks (New rows, Total infringements ...) ab Settings & Generate me nahi, Process log section me aate hain.

v1.15.9 changes
- Header centred + "Created by Chandra Prakash Tripathi" (right). Section headers centred. 3 top cards ka purana size wapas, sirf colours. Settings & Generate broader (inputs row + 3 bade buttons). Shadows stronger.
- Generate ke baad ka log ab alag "Process log" section me square blocks me aata hai.

v1.15.8 changes
- UI: white background, fit-to-screen (no page scroll), colour-contrast header/cards/buttons, balanced card + button sizes.
- Select titles: ab files upload karne par sabhi titles by default UNselected rehte hain (user khud select kare).

v1.15.7 changes
- UI: Neumorphism style (neu.css) + slide motion - cards/buttons/popups slide in, drawer slides from right, buttons lift on hover and press in on click, process line fills as steps finish.

v1.15.6 changes
- ORDER CHANGE (Live + Comet): pehle New data append hota hai, phir sabse last me Removal update lagta hai. Isliye removal file me naye data ke URL bhi hon to wo bhi Completed ho jaate hain.

v1.15.5 changes
- YouTube/X/Facebook/Instagram: Removal Time blank ho (ya Removal Status Dead na ho) to "Active URLs" popup aata hai.

v1.15.4 changes
- Summary: "Total infringements/takedowns since prior report (insert time)" me text ab "EST" (time same, sirf IST label -> EST).
- Report name (file name) bharna zaroori: jab tak naam na do, teeno Generate buttons disabled rehte hain.

v1.15.3 changes
- "Report time (IST)" option hata diya - report time ab hamesha abhi ka time (IST) rehta hai. File ka naam "Report name (file name)" box se aata hai (blank = purana auto naam).
- NOTES TAB FIX: previous report ke Notes tab ka text box ab output me as-is aata hai (pehle ExcelJS use hata deta tha, isliye blank aata tha). Live aur Comet dono me. Naya notes-core.js + lib/jszip.min.js.

v1.15.2 changes
- NEW popup: new data me Enforcement time missing ya Dolby status blank (data file me 'dolby' naam ka column ho aur khali ho) wale rows mile to list dikhti hai. 'Still want to generate' par hi report banti hai; Cancel = report nahi banti. Live, Comet, Findings teeno me.

v1.15.1 changes
- EXCEL CORRUPTION FIX: har export se pehle workbook clean hota hai (NaN/Infinity numbers, invalid dates, control characters, 32767+ char strings) - inse Excel "found a problem / recovery" dikhata tha aur data gadbad hota tha.
- "Total takedowns since prior report (insert time)": removal file upload na ho to ab hamesha 00 (purani report ka number nahi rehta).
- File size: "GBs/MBs/KBs/TBs" -> "GB/MB/KB/TB" (Live, Comet, Findings sab me, previous report ke purane rows me bhi).

v1.15.0 changes
- NEW 3rd button "Generate Findings Report": previous report ki zaroorat nahi, sirf New data files (+ optional removal file) se 1-sheet "Findings" report banti hai (logo, 16 columns, EST) - sample Audible_Findings jaisa. Completed = data me Removal Status Dead ya removal file me URL; baaki N. Naam auto: MMDDYY_Findings_<since day>_<Month>.xlsx (ya Report name box wala).
- "Report name (optional)" box: bhar do to Live/Comet dono report isi naam se download hoti hain (.xlsx khud lagta hai, galat characters hat jaate hain). Khali = purana auto naam.
- FILE FORMAT CHECK: har Browse / drag-drop par file ka format check hota hai. Galat ho to error dialog aata hai (file reason ke saath) aur file add nahi hoti.
  Check: extension (.xlsx / .xls) + asli Excel content (renamed CSV/txt pakdi jaati hai) + andar ka structure:
  Previous Live Report = Findings / Takedown Pending / Takedown Completed tabs (URL column ke saath); Previous Comet Report = "List" sheet ke saare Comet columns;
  Removal file = URL + Removal Time columns (aur kam se kam 1 URL); New data = platform button ke hisaab se (Facebook button me Linking file -> error, etc.).
  Live file Comet picker me (ya ulta) daalo to error me bata deta hai ki ye kaun si report lag rahi hai.
- COMET: jin URLs ka Takedown Completed = Y hai, unka "Reason Pending" = "-". Removal se naye complete hue rows me bhi, aur previous Comet file ke purane completed rows me bhi.
- manifest.json zip me khali (0 bytes) tha - ek chalne wala Manifest V3 manifest add kiya. Apna original manifest ho to wahi rakho.

v1.14.0 changes
- Do buttons: "Generate Live Report" (pehle jaisa) aur "Generate Comet Report" (naya). Card 1 me Previous Live Report aur Previous Comet Report ke alag pickers.
- Comet report = previous Comet file ("List" sheet, 32 columns) me naye URL neeche append + removal file wale URL complete. IST -> EST = minus 10h 30m (Identification EST, Takedown Initiated EST, Takedown Completion EST).
  Naye URL: Takedown Initiated Y, Pending Y, Completed N, Dolby blank, Counter Notice N. Removal: Pending N, Completed Y, completion date/time (EST), TAT Hours (1 ghante se kam = 0:MM).
- File Size (naye platforms): raw file me size ho to wahi, warna "Streaming" (Signpost/Full/Clip sab); Market Place -> "Paid". Market Place Channel/Domain = base domain (ebay.com).
- Social Media & UGC removal ab Major Social + Takedown Completed dono me jata hai (existing report jaisa).

v1.13.0 changes
- NEW data upload: koi active/dead check nahi. Har naya URL sirf Findings + Takedown Pending (+ Delisting, Linking ke liye) me jata hai. Takedown Completed / Major Social Platform Take down NAHI badalte.
  Summary: Total Infringements, Total infringements since prior report, Open issues remaining, Current Takedown Percentage.
- REMOVAL file upload: tabhi Takedown Completed + Major Social Platform Take down update hote hain.
  Summary: Total Takedowns, Total takedowns since prior report, Open issues remaining, Current Takedown Percentage.
- Market Place ke liye active-URL popup nahi aata (Facebook/Instagram/YouTube/X ke liye aata hai).
- Facebook/Instagram/YouTube/X ke naye Pending rows me "Reason Takedown is Pending:" blank.

v1.12.0 changes
- Instagram aur Market Place (eBay etc.) ka merge add (Facebook jaisa: Social + Takedown Completed dono, active URL popup).
- New data files card: platform ke naam ke horizontal buttons - har button hi uska Browse button hai (click = file chuno, drag-drop bhi chalta hai). Upload ke baad button me rows count + ✓ dikhta hai, ✕ se hatao. "+ Other..." se naya platform.

