Attribute VB_Name = "CarVendors"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Public Sub CarManufacturers(c1 As ComboBox)
c1.AddItem ("Maruti Udyog Ltd")
c1.AddItem ("Honda Siel Cars")
c1.AddItem ("Hyundai Motor")
c1.AddItem ("Toyota Kirloskar Motor")
c1.AddItem ("Skoda")
c1.AddItem ("Tata Motors")
c1.AddItem ("Mahindra And Mahindra")
c1.AddItem ("General Motors")
c1.AddItem ("Hindustan Motors")
c1.AddItem ("DaimlerChrysler")
c1.AddItem ("Ford Motors")
c1.AddItem ("Nissan Motors")
c1.AddItem ("Fiat")
c1.AddItem ("Porsche")
c1.AddItem ("Ferrari")
c1.AddItem ("Mitsubishi")
c1.AddItem ("Dodge")
c1.AddItem ("B M W")
End Sub
Public Sub Carmakes(MAKE As ComboBox, Manufacturer As String)
MAKE.CLEAR

Select Case Manufacturer

Case "Maruti Udyog Ltd"
MAKE.AddItem ("Omni")
MAKE.AddItem ("800")
MAKE.AddItem ("Gypsy")
MAKE.AddItem ("Zen")
MAKE.AddItem ("Alto")
MAKE.AddItem ("WagonR")
MAKE.AddItem ("Esteem")
MAKE.AddItem ("Baleno")
MAKE.AddItem ("Vitara")
MAKE.AddItem ("Versa")

Case "Honda Siel Cars"
MAKE.AddItem ("City")
MAKE.AddItem ("Accord")
MAKE.AddItem ("CRV")
MAKE.AddItem ("NSX")

Case "Hyundai Motor"
MAKE.AddItem ("Santro")
MAKE.AddItem ("GetZ")
MAKE.AddItem ("Accent")
MAKE.AddItem ("Sonata")
MAKE.AddItem ("Elantra")
MAKE.AddItem ("Terracan")
MAKE.AddItem ("Starex")
MAKE.AddItem ("Tiburon")

Case "Toyota Kirloskar Motor"
MAKE.AddItem ("Corolla")
MAKE.AddItem ("Innova")
MAKE.AddItem ("Camry")
MAKE.AddItem ("Qualis")
MAKE.AddItem ("Land Cruiser Prado")
MAKE.AddItem ("Aygo")
MAKE.AddItem ("Peugeot")
MAKE.AddItem ("Citroen")

Case "Skoda Auto"
MAKE.AddItem ("Octavia")
MAKE.AddItem ("Superb")

Case "Tata Motors"
MAKE.AddItem ("Indica")
MAKE.AddItem ("Indigo")
MAKE.AddItem ("Sumo")
MAKE.AddItem ("Safari")
MAKE.AddItem ("Seirra")

Case "Mahindra And Mahindra"
MAKE.AddItem ("Scorpio")
MAKE.AddItem ("Bolero")

Case "General Motors"
MAKE.AddItem ("Opel")
MAKE.AddItem ("Chevrolet")

Case "Hindustan Motors"
MAKE.AddItem ("Ambassador")

Case "DaimlerChrysler"
MAKE.AddItem ("Mercedes Benz")
MAKE.AddItem ("Maybach")
MAKE.AddItem ("Sebring")
MAKE.AddItem ("Crossfire")

Case "Fiat"
MAKE.AddItem ("Ferrari")
MAKE.AddItem ("Palio")
MAKE.AddItem ("Uno")
MAKE.AddItem ("Croma")
MAKE.AddItem ("Petra")


Case "Ford Motors"
MAKE.AddItem ("Ikon")
MAKE.AddItem ("Escort")
MAKE.AddItem ("Martini")
MAKE.AddItem ("Fusion")
MAKE.AddItem ("Endeavor")
MAKE.AddItem ("Puma")
MAKE.AddItem ("GT40")

Case "Nissan Motors"
MAKE.AddItem ("X-Trail")
MAKE.AddItem ("Zaroot")

Case "Mitsubishi"
    MAKE.AddItem ("Lancer")
    MAKE.AddItem ("GT")
    MAKE.AddItem ("Pajero")
    

End Select
End Sub
Public Sub Carmodels(MODEL As ComboBox, MAKE As String)
MODEL.CLEAR

Select Case MAKE
'MUL
Case "Omni"
    MODEL.AddItem ("Omni")
Case "800"
    MODEL.AddItem ("800")
    MODEL.AddItem ("800 DLX")
    MODEL.AddItem ("800 AC")
Case "Gypsy"
    MODEL.AddItem ("King Hard Top (Metallic)")
    MODEL.AddItem ("King Soft Top (Metallic)")
Case "Zen"
    MODEL.AddItem ("Zen LX ")
    MODEL.AddItem ("Zen VX")
Case "Alto"
    MODEL.AddItem ("Alto")
    MODEL.AddItem ("Alto LX")
    MODEL.AddItem ("Alto Spin")
Case "WagonR"
    MODEL.AddItem ("WagonR")
Case "Esteem"
    MODEL.AddItem ("Esteem")
    MODEL.AddItem ("Esteem LXI")
    MODEL.AddItem ("Esteem DI")
    MODEL.AddItem ("Esteem VX PS")
Case "Baleno"
    MODEL.AddItem ("Baleno")
Case "Vitara"
    MODEL.AddItem ("Vitara")
    MODEL.AddItem ("Vitara XL7")
Case "Versa"
    MODEL.AddItem ("Versa")

'Honda Siel Cars
Case "City"
    MODEL.AddItem ("City")
Case "Accord"
    MODEL.AddItem ("Accord")
    MODEL.AddItem ("Accord 2.4")
    MODEL.AddItem ("Accord V6")
Case "CRV"
    MODEL.AddItem ("CV-R")
    MODEL.AddItem ("CV-R MT")
    MODEL.AddItem ("CV-R AT")
Case "NSX"
    MODEL.AddItem ("NSX")
    
'Hyundai Motor
Case "Santro"
    MODEL.AddItem ("Santro")
    MODEL.AddItem ("Santro Xing")
    MODEL.AddItem ("Santro Xing 1.1")
Case "GetZ"
    MODEL.AddItem ("GetZ")
Case "Accent"
    MODEL.AddItem ("Accent")
    MODEL.AddItem ("Accent Viva 1.6")
    MODEL.AddItem ("Accent 1.5 CRDi")
Case "Sonata"
    MODEL.AddItem ("Sonata")
Case "Elantra"
    MODEL.AddItem ("Elantra")
    MODEL.AddItem ("Elantra 1.8 GLS")
    MODEL.AddItem ("Elantra CRDi")
    MODEL.AddItem ("Elantra 2.0 CRDi")
Case "Terracan"
    MODEL.AddItem ("Terracan")
Case "Starex"
    MODEL.AddItem ("Starex")
Case "Tiburon"
    MODEL.AddItem ("Tiburon")
    
'Toyota Kirloskar Motor
Case "Corolla"
    MODEL.AddItem ("Corolla")
Case "Innova"
    MODEL.AddItem ("Innova")
Case "Camry"
    MODEL.AddItem ("Camry")
Case "Qualis"
    MODEL.AddItem ("Qualis")
Case "Land Cruiser Prado"
    MODEL.AddItem ("Land Cruiser Prado")
Case "Aygo"
    MODEL.AddItem ("Aygo")
Case "Peugeot"
    MODEL.AddItem ("Peugeot")
Case "Citroen"
    MODEL.AddItem ("Citroen")
    MODEL.AddItem ("Citroen C1")
    MODEL.AddItem ("Citroen C5 SX")
    MODEL.AddItem ("Citroen C5 Exclusive")
    
'Skoda Auto
Case "Octavia"
    MODEL.AddItem ("Octivia")
    MODEL.AddItem ("Octivia 2.0 Petrol")
    MODEL.AddItem ("Octivia 1.8T RS")
    MODEL.AddItem ("Octivia TDi")
    MODEL.AddItem ("Octivia 1.9 TDi")
    MODEL.AddItem ("Octivia TDi Auto")
    MODEL.AddItem ("Octivia Rider")
    MODEL.AddItem ("Octivia Ambiente")
Case "Superb"
    MODEL.AddItem ("Superb")

'Tata Motors
Case "Indica"
    MODEL.AddItem ("Indica")
    MODEL.AddItem ("Indica V2")
Case "Indigo"
    MODEL.AddItem ("Indigo")
    MODEL.AddItem ("Indigo Marina")
Case "Sumo"
    MODEL.AddItem ("Sumo Victa Turbo")
    MODEL.AddItem ("Sumo Victa Spacio")
    MODEL.AddItem ("Sumo SE+Euro I")
    MODEL.AddItem ("Sumo LX 4 x 2")
Case "Safari"
    MODEL.AddItem ("Safari")
Case "Seirra"
    MODEL.AddItem ("Seirra")
    c2.AddItem ("Scorpio TC-2 4WD Turbo 2.6")

'Mahindra And Mahindra
Case "Scorpio"
    MODEL.AddItem ("Scorpio 2.6 Turbo")
    MODEL.AddItem ("Scorpio REV 116")
    MODEL.AddItem ("Scorpio CRDe")
    MODEL.AddItem ("Scorpio TC-4")
    MODEL.AddItem ("Scorpio SLX")
    MODEL.AddItem ("Scorpio Euro II")
Case "Bolero"
    MODEL.AddItem ("Bolero")
    MODEL.AddItem ("Bolero Invader")
    MODEL.AddItem ("Bolero XLS")
    MODEL.AddItem ("Bolero GLX")
    MODEL.AddItem ("Bolero XDB")
    MODEL.AddItem ("Bolero 4WD")
    MODEL.AddItem ("Bolero BS2")
    MODEL.AddItem ("Bolero Camper")
    MODEL.AddItem ("Bolero SportZ")
    MODEL.AddItem ("Bolero 2WD")
    MODEL.AddItem ("Bolero 7S")
    MODEL.AddItem ("Bolero 10S")
    MODEL.AddItem ("Bolero PS&AC")
    
'General Motors
Case "Opel"
    MODEL.AddItem ("Opel Astra")
    MODEL.AddItem ("Opel Vectra")
    MODEL.AddItem ("Opel Corsa")
    MODEL.AddItem ("Opel Corsa 1.4")
    MODEL.AddItem ("Opel Corsa 1.6")
    MODEL.AddItem ("Opel Corsa Swing 1.6")
    MODEL.AddItem ("Opel Corsa SAIL")
    MODEL.AddItem ("Opel Corsa SAIL 1.4")
    MODEL.AddItem ("Opel Corsa SAIL 1.6")
    
Case "Chevrolet"
    MODEL.AddItem ("Tavera")
    MODEL.AddItem ("Tavera B3-10 BS II Metallic")
    MODEL.AddItem ("Tavera D1-8 BS II Metallic")
    MODEL.AddItem ("Optra")
    MODEL.AddItem ("Optra 1.6")
    MODEL.AddItem ("Optra 1.8")
    MODEL.AddItem ("Forester")
    MODEL.AddItem ("Spark")
    MODEL.AddItem ("Corvette")


'Hindustan Motors
Case "Ambassador"
    MODEL.AddItem ("Ambassador")

'DaimlerChrysler
Case "Mercedes Benz"
    MODEL.AddItem ("E 270 CDI")
    MODEL.AddItem ("M Class")
    MODEL.AddItem ("B Class")
    MODEL.AddItem ("C 200 Kompressor")
    MODEL.AddItem ("C 220 CDI")
    MODEL.AddItem ("E 200 K")
    MODEL.AddItem ("E 240")
    MODEL.AddItem ("E 270 CDI")
    
Case "Maybach"
    MODEL.AddItem ("Maybach")
Case "Sebring"
    MODEL.AddItem ("Sebring")
Case "Crossfire"
    MODEL.AddItem ("Crossfire")

'Fiat
Case "Ferrari"
    MODEL.AddItem ("F250 GT SWB California Spider")
    MODEL.AddItem ("F360")
    MODEL.AddItem ("F355 Spider")
    MODEL.AddItem ("F360 Spider")
    MODEL.AddItem ("F430 Spider")
    MODEL.AddItem ("Daytona")
    
Case "Palio"
    MODEL.AddItem ("Palio")
    MODEL.AddItem ("Palio NV")
    MODEL.AddItem ("Palio 1.9D")
Case "Uno"
    MODEL.AddItem ("Uno")
Case "Croma"
    MODEL.AddItem ("Croma")
Case "Petra"
    MODEL.AddItem ("Petra 1.6")
    MODEL.AddItem ("Petra 1.9D")

'Ford Motors
Case "Ikon"
    MODEL.AddItem ("Ikon")
    MODEL.AddItem ("Ikon 1.6")
    MODEL.AddItem ("Ikon 1.3")
    MODEL.AddItem ("Ikon D")
    MODEL.AddItem ("Ikon 1.8D")
    
Case "Escort"
    MODEL.AddItem ("Escort")
Case "Martini"
    MODEL.AddItem ("Martini")
Case "Fusion"
    MODEL.AddItem ("Fusion")
    MODEL.AddItem ("Fusion Plus")
    MODEL.AddItem ("Fusion Plus ABS")
Case "Endeavor"
    MODEL.AddItem ("Endeavor")
    MODEL.AddItem ("Endeavour 4 x 2")
    MODEL.AddItem ("Endeavour 4 x 4")
Case "Puma"
    MODEL.AddItem ("Puma")
Case "GT40"
    MODEL.AddItem ("GT40")

'Nissan Motors
Case "X-Trail"
    MODEL.AddItem ("X-Trail")
    MODEL.AddItem ("X-Trail - Comfort")
    MODEL.AddItem ("X-Trail - Elegance")
Case "Zaroot"
    MODEL.AddItem ("Zaroot")
    
'mitsubishi
Case "Lancer"
    MODEL.AddItem ("Lancer")
    MODEL.AddItem ("Lancer LX 1.8")
    MODEL.AddItem ("Lancer LX 2000 Disel")
    MODEL.AddItem ("Lancel Evolution V")
Case "GT"
    MODEL.AddItem ("GT 3000")
Case "Pajero"
    MODEL.AddItem ("Pajero")
    MODEL.AddItem ("Pajero 2.8")

   
End Select
End Sub


Public Sub Load_CarVendors(c1 As ComboBox)
'vendors of cars on indian roads
c1.AddItem ("Maruti")
c1.AddItem ("Honda")
c1.AddItem ("Hyundai")
c1.AddItem ("Toyota")
c1.AddItem ("Skoda")
c1.AddItem ("Ford")
c1.AddItem ("Tata")
c1.AddItem ("Mahindra")
c1.AddItem ("Daewoo")
c1.AddItem ("GM Opel")
c1.AddItem ("GM Chevrolet")
c1.AddItem ("Nissan")
c1.AddItem ("Fiat")

'SuperCar Vendors
c1.AddItem ("Dodge")
c1.AddItem ("Aston Martin")
c1.AddItem ("Lamborghini")
c1.AddItem ("Porsche")
c1.AddItem ("Bugatti")
c1.AddItem ("Jaguar")
c1.AddItem ("Ferrari")
c1.AddItem ("McLaren")
c1.AddItem ("Mercedes Benz")
c1.AddItem ("BMW")
c1.AddItem ("Suzuki")
c1.AddItem ("Infiniti")
c1.AddItem ("Peugeot")
c1.AddItem ("Lotus")
c1.AddItem ("Audi")
c1.AddItem ("Mini")
c1.AddItem ("Lexus")
c1.AddItem ("Milano")
c1.AddItem ("Citroen")
c1.AddItem ("Bentley")
c1.AddItem ("Daimler Chrysler")
End Sub

Public Sub Load_CarModels(c2 As ComboBox, VENDOR As String)
c2.CLEAR

Select Case VENDOR

Case "Maruti"
c2.AddItem ("Omni")
c2.AddItem ("800")
c2.AddItem ("800 DLX")
c2.AddItem ("Suzuki Gypsy King Soft Top (Metallic)")
c2.AddItem ("Suzuki Gypsy King Hard Top (Metallic)")
c2.AddItem ("Zen")
c2.AddItem ("Zen LX")
c2.AddItem ("Zen VX")
c2.AddItem ("Alto")
c2.AddItem ("Alto Spin")
c2.AddItem ("WagonR")
c2.AddItem ("Esteem")
c2.AddItem ("Esteem LXi")
c2.AddItem ("Esteem Di")
c2.AddItem ("Esteem VX PS")
c2.AddItem ("Baleno")
c2.AddItem ("Vitara-XL7")

Case "Hyundai"
c2.AddItem ("Santro")
c2.AddItem ("Santro Xing")
c2.AddItem ("GetZ")
c2.AddItem ("Accent")
c2.AddItem ("Sonata")
c2.AddItem ("Elantra")
c2.AddItem ("Elantra CRDi")
c2.AddItem ("Terracan")
c2.AddItem ("Starex")
c2.AddItem ("Tiburon")

Case "Ford"
c2.AddItem ("Ikon")
c2.AddItem ("Ikon D")
c2.AddItem ("Escort")
c2.AddItem ("Martini")
c2.AddItem ("Fusion")
c2.AddItem ("Fusion +")
c2.AddItem ("Fusion +(ABS)")
c2.AddItem ("Endeavor")
c2.AddItem ("Endeavour 4 x 2")
c2.AddItem ("Endeavour 4 x 4")
c2.AddItem ("Puma")
c2.AddItem ("GT40 ")

Case "Fiat"
c2.AddItem ("Palio")
c2.AddItem ("Croma")

Case "Nissan"
c2.AddItem ("X-Trail")
c2.AddItem ("X-Trail - Comfort")
c2.AddItem ("X-Trail - Elegance")
c2.AddItem ("Zaroot")

Case "Tata"
c2.AddItem ("Indica")
c2.AddItem ("Indica V2")
c2.AddItem ("Indigo")
c2.AddItem ("Indigo Marina")
c2.AddItem ("Sumo")
c2.AddItem ("Safari")
c2.AddItem ("Sumo SE+Euro I")
c2.AddItem ("Safari LX 4 x 2")
c2.AddItem ("Seirra")

Case "Daewoo"
c2.AddItem ("Matiz")
c2.AddItem ("Cello")

Case "Mahindra"
c2.AddItem ("Scorpio TC-2 4WD Turbo 2.6")
c2.AddItem ("Scorpio TC-4 Turbo 2.6 SLX Euro II")
c2.AddItem ("Bolero Invader GLX XDB 4WD BS2")
c2.AddItem ("Bolero Camper 4WD BS2")
c2.AddItem ("Bolero Sportz 2WD BS2 7S (PS&AC)")
c2.AddItem ("Bolero XLS 2WD BS2 10S (PS&AC)")

Case "Honda"
c2.AddItem ("City")
c2.AddItem ("Accord")
c2.AddItem ("CR-V")
c2.AddItem ("CRV MT")
c2.AddItem ("CRV AT")
c2.AddItem ("NSX")

Case "Toyota"
c2.AddItem ("Corolla")
c2.AddItem ("Innova")
c2.AddItem ("Camry")
c2.AddItem ("Qualis")
c2.AddItem ("Land Cruiser Prado")
c2.AddItem ("Aygo")
c2.AddItem ("Peugeot")
c2.AddItem ("Citroen C1")

Case "Aston Martin"
c2.AddItem ("AmV8 Vantage")
c2.AddItem ("Vanquish")
c2.AddItem ("DB7")

Case "Audi"
c2.AddItem ("A8")
c2.AddItem ("A6")
c2.AddItem ("TT Coupe")
c2.AddItem ("RS4 Avant")


Case "Mercedes Benz"
c2.AddItem ("E270 CDI")
c2.AddItem ("M Class")
c2.AddItem ("B Class")
c2.AddItem ("")

Case "Bentley"
c2.AddItem ("Arnage")
c2.AddItem ("Continental GT")
c2.AddItem ("")

Case "Porsche"
c2.AddItem ("Carrera GT")
c2.AddItem ("Cayenne")
c2.AddItem ("Turbo")
c2.AddItem ("Boxster ")

Case "Lamborghini"
c2.AddItem ("Gallardo")
c2.AddItem ("Diablo 30SE")
c2.AddItem ("Diablo")
c2.AddItem ("SL VSi_Xi")

Case "Dodge"
c2.AddItem ("Viper")
c2.AddItem ("Viper GTS")
c2.AddItem ("Viper GTSR")

Case "Bugatti"
c2.AddItem ("EB 110")
c2.AddItem ("Eb 16.4 Veyron")

Case "BMW"
c2.AddItem ("3 Serie")
c2.AddItem ("3 Serie Coupe")
c2.AddItem ("Xcoupe")
c2.AddItem ("323i")
c2.AddItem ("328Ci")
c2.AddItem ("328i")
c2.AddItem ("528i")
c2.AddItem ("540i")
c2.AddItem ("Z301")
c2.AddItem ("Z801")
c2.AddItem ("Z802")
c2.AddItem ("M501")
c2.AddItem ("850Ci")
c2.AddItem ("Z3 Roadster")
c2.AddItem ("540i sedan")
c2.AddItem ("325i sport wagon")
c2.AddItem ("330Ci convertible")
c2.AddItem ("330xi sedan")
c2.AddItem ("X5 4.4i SAV")
c2.AddItem ("323t1 compact")

Case "Mini"
c2.AddItem ("Coopers")

Case "Lexus"
c2.AddItem ("GS 400")

Case "Milano"
c2.AddItem ("Alfa Romeo")
c2.AddItem ("Alfa Romeo Spider")

Case "Citroen"
c2.AddItem ("C5 Exclusive")
c2.AddItem ("c5 SX")

Case "Mitsubishi"
c2.AddItem ("R22 MRE")
c2.AddItem ("Lancer")
c2.AddItem ("Lancel Evolution V")
c2.AddItem ("GT 3000")
c2.AddItem ("Pajero")

Case "GM Opel"
c2.AddItem ("Astra")
c2.AddItem ("Vectra")
c2.AddItem ("Corsa")
c2.AddItem ("Corsa SAIL")

Case "GM Chevrolet"
c2.AddItem ("Tavera")
c2.AddItem ("Tavera B3-10 BS II Metallic")
c2.AddItem ("Tavera D1-8 BS II Metallic")
c2.AddItem ("Optra")
c2.AddItem ("Optra 1.6")
c2.AddItem ("Forester")
c2.AddItem ("Spark")
c2.AddItem ("Corvette")

Case "Skoda"
c2.AddItem ("Octavia")
c2.AddItem ("Octavia TDi")
c2.AddItem ("Octavia Rider")
c2.AddItem ("Octavia Ambiente")
c2.AddItem ("RS")
c2.AddItem ("Superb")

Case "Peugeot"
c2.AddItem ("206 CC")

Case "Lotus"
c2.AddItem ("Esprint V8")
c2.AddItem ("Esprint GT3")
c2.AddItem ("Elise")

Case "Suzuki"
c2.AddItem ("Grand Vitara XL.7")
c2.AddItem ("Swift")


Case "Ferrari"
c2.AddItem ("F250 GT SWB California Spider")
c2.AddItem ("F360")
c2.AddItem ("F355 Spider")
c2.AddItem ("F360 Spider")
c2.AddItem ("F430 Spider")
c2.AddItem ("Daytona")


Case "Daimler Chrysler"
c2.AddItem ("Maybach")
c2.AddItem ("Sebring")
c2.AddItem ("Stratus")
c2.AddItem ("300 M")
c2.AddItem ("Crossfire")

Case "Jaguar"
c2.AddItem ("XKR")

End Select
End Sub
