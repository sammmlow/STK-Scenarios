# STK Master Integration Certification Script

# Full Script Based on Jupyter Notebook
# Tutorial script, from AGI, editted by Samuel Low, 11/08/2020.
##############################################################################
##############################################################################

#Import basic utilities
import datetime as dt
import numpy as np
import os

#Needed to interact with COM
from comtypes.client import CreateObject

##############################################################################
##############################################################################

# Start STK10 Application

# uiApp is the pointer to the IAgUiApplication interface used to make the STK
# application window visible and user-controlled.

# In computer science, a pointer is an object in many programming languages
# that stores a memory address. This can be that of another value located in
# computer memory, or in some cases, that of memory-mapped computer hardware.
# A pointer references a location in memory, and obtaining the value stored at
# that location is known as dereferencing the pointer. As an analogy, a page
# number in a book's index could be considered a pointer to the corresponding
# page; dereferencing such a pointer would be done by flipping to the page
# with the given page number and reading the text found on that page. The
# actual format and content of a pointer variable is dependent on the
# underlying computer architecture.

# COM defines a COM object with a set of data and related functions. You use
# the set of related functions to access the COM object's data.

# These sets of functions are called interfaces. The individual functions of
# an interface are called methods. COM uses pointers to gain access to the
# methods of an interface.

# uiApp: <class 'comtypes.POINTER(_IAgUiApplication)'>
uiApp = CreateObject("STK10.Application") 

##############################################################################
##############################################################################

# Gets helpful info on classes and methods of the IAgUiApplication interface
# help(uiApp) # You must close the UI first

##############################################################################
##############################################################################

uiApp.Visible = True

#Set the application so that it is user controlled.
uiApp.UserControl = True

##############################################################################
##############################################################################

# This line of code gets the pointer to the IAgiStkObjectRoot interface.
stkRoot = uiApp.Personality2

##############################################################################
##############################################################################
print("\ntype(stkRoot):")
print(type(stkRoot))

# At this point it is important to highlight the difference between the object
# AgStkObjectRoot and the interface IAgStkObjectRoot. Ag is used for objects
# and IAg for interfaces. As mentioned before, a COM object holds data that
# you can only access through sets of related functions. These sets of
# functions are called interfaces and the individual functions are called
# methods. The only way to access an interface's method is through a pointer
# to the interface. A single COM interface does not necessarily hold all the
# functions associated with a the COM object. For example, theIAgStkObjectRoot
# interface contains most of the functions for the AgStkObjectRoot object,
# but not all the functions. Finally, root.Personality2 returns a pointer to
# the IAgStkObjectRoot interface and not the AgStkObjectRoot object itself.


##############################################################################
##############################################################################

# When you execute stkRoot=app.Personality2 for the first time on your
# computer, you generate the STKObjects, STKUtil libraries, and several other
# comtypes libraries. You can find all the comptypes libraries you generated
# in the comtypes gen folder. After running stkRoot=app.Personality2 at least
# once on your computer, you can run the following two lines of code at any time
# to import the newly generated STKObject and STKUtil libraries.

from comtypes.gen import STKObjects
from comtypes.gen import STKUtil

##############################################################################
##############################################################################

# When you executed uiApp = CreateObject("STK10.Application"), you placed
# other libraries into the comtypes gen folder as well. You should know where
# all the Python generated STK Object Model libraries are stored, because
# occasionally you may need to remove them manually. For example, if you
# upgrade from STK 11 to STK 12, you would have to remove these generated
# libraries. To locate the directory where the comptypes are stored, run:

from comtypes.client import gen_dir
print("\ngen_dir: ")
print(gen_dir)
print("\ngen_dir contents:")
print(os.listdir(gen_dir))

##############################################################################
##############################################################################

# Create a new scenario using the NewScenario method of the IAgStkObjectRoot
# interface. The NewScenario method expects an input string as the name of the
# scenario but it does not return anything.

stkRoot.NewScenario("IntegrationCertification")

# Get a reference to the scenario object (null if no scenario has been loaded)
scenario = stkRoot.CurrentScenario

##############################################################################
##############################################################################

print("\ntype(scenario):")
print(type(scenario))

##############################################################################
##############################################################################

# Investigating scenario with the Python built-in dir() lists out the
# available properties and methods of scenario. Unfortunately, you cannot
# access the IAgScenario interface methods and properties through the
# pointer to IAgStkObject.

print("\ndir(scenario):")
print(dir(scenario))

##############################################################################
##############################################################################

# The QueryInterface() Method: Fortunately, one interface can give you
# pointers for the other interfaces within the same object. This means you
# can obtain a pointer to IAgScenario from the pointer to IAgStkObject.
# The QueryInterface() method returns a pointer to the desired interface.
# If the inteface is not implemented then the QueryInterface() method raises
# an exception.

# Some more notes on QueryInterface by Microsoft documentation on COM...
# Every interface is derived from IUnknown, so every interface has an
# implementation of QueryInterface. Regardless of implementation, this
# method queries an object using the IID of the interface to which the
# caller wants a pointer. If the object supports that interface,
# QueryInterface retrieves a pointer to the interface, while also
# calling AddRef. Otherwise, it returns the E_NOINTERFACE error code.

scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)

print("\ntype(scenario2):")
print(type(scenario2))

scenario2.StartTime = "1 Jun 2016 15:00:00.000"

scenario2.StopTime = "2 Jun 2016 15:00:00.000"

#Reset STK to the new start time
stkRoot.Rewind()

import time
print("Adding a 20 second delay...")
time.sleep(10)
print("10 more seconds")
time.sleep(10)

##############################################################################
##############################################################################

# Example of first 3 lines of the Facilities.txt file
# Fac01,-158.26,21.57
# Fac02,-147.52,64.98
# Fac03,-120.60,34.67

# This means that
# facilityData[0] is the facility name
# facilityData[1] is the longitude in degrees
# facilityData[2] is the latitude in degrees.


# You can send these Connect commands with the ExecuteCommand method, which
# is part of the IAgStkObjectRoot interface. The Connect Command Listings may
# be useful when completing this task. 

with open("Facilities.txt", "r") as facilityFile:
    for line in facilityFile:
        facilityData = line.strip().split(",")
        
        
        insertNewFacCmd = f"New / */Facility {facilityData[0]}"
        stkRoot.ExecuteCommand(insertNewFacCmd)
        
        # setPositionCmd = cmd + path + coord + p_lat + p_lon + p_alt
        setPositionCmd  = f"SetPosition */Facility/{facilityData[0]} "
        setPositionCmd += "Geodetic " + facilityData[2] + " "
        setPositionCmd += facilityData[1] + " 0.0"
        
        stkRoot.ExecuteCommand(setPositionCmd)
        
        # Set the colour of the marker
        setColorCmd = f"Graphics */Facility/{facilityData[0]} SetColor cyan"
        stkRoot.ExecuteCommand(setColorCmd)
        
facilityFile.close()

# ##############################################################################
# ##############################################################################

# # You can either pass the object member STKObjects.eSatellite, or you can
# # use the enumerator index value of 18.
# satellite = scenario.Children.New(STKObjects.eSatellite, "TestSatellite")

# # Cast the satellite to obtain a pointer to IAgSatellite
# satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite)

# ##############################################################################
# ##############################################################################

# # The IAgSatellite interface exposes the Graphics property, which gives a
# # pointer to the IAgSaGraphics interface. The IAgSaGraphics interface exposes
# # the Attributes property, which gives a pointer to the IAgVeGfxAttributes
# # interface. The IAgVeGfxAttributes interface is a base interface for vehicle
# # 2D graphics attributes and does not expose any methods. This might seem like
# # a dead end, but remember that the IAgVeGfxAttributes interface is implemented
# # by an object that may implement more than one interface, and the
# # IAgVeGfxAttributes interface may not represent all of the methods the object
# # supports. AgVeGfxAttributesOrbit is the class that implements the interface,
# # and represents the 2D graphics attributes of a satellite. The
# # IAgVeGfxAttributesBasic and IAgVeGfxAttributesOrbit interfaces can make the
# # desired changes to the satellite’s 2D Graphics Attributes.

# #Change some basic display attributes
# satelliteBasicGfxAttributes =  satellite2.Graphics.Attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)
# satelliteBasicGfxAttributes.Color = 65535 #Yellow
# satelliteBasicGfxAttributes.Line.Width = STKObjects.e2     

# # Set inheritance of 2D graphics settings from the scenario level to false.
# satelliteBasicGfxAttributes.Inherit = False

# satelliteBasicGfxAttributes.QueryInterface(STKObjects.IAgVeGfxAttributesOrbit).IsGroundTrackVisible = False

# ##############################################################################
# ##############################################################################

# # Select Propagator
# satellite2.SetPropagatorType(STKObjects.ePropagatorTwoBody)

# ##############################################################################
# ##############################################################################

# # Next you need to define the satellite's initial state with the following
# # orbital elements: semimajor axis, eccentricity, inclination, argument of
# # perigee, RAAN, and true anomaly. To accomplish this, you must build an
# # initial state, assign it to the satellite, and finally propagate the orbit.

# # You can gain access to the initial orbit state through the satellite's
# # propagator object. In the block below, get a pointer to the
# # IAgVePropagtorTwoBody interface. Then use that pointer to convert the obit
# # state into the classical representation, and obtain a pointer to the
# # IAgOrbitStateClassical interface.

# # Get the Two Body Propagator interface
# twoBodyPropagator = satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)

# keplarian = twoBodyPropagator.InitialState.Representation.ConvertTo(STKUtil.eOrbitStateClassical).QueryInterface(STKObjects.IAgOrbitStateClassical)

# # With the IAgOrbitStateClassical interface you will be able to set the values
# # of the desired orbital elements.

# ##############################################################################
# ##############################################################################

# # The SizeShape property only provides a pointer to the IAgClassicalSizeShape
# # interface, which does not immediately provide access to the semimajor axis
# # or eccentricity values. To access those, you "cast" to the
# # IAgClassicalSizeShapeSemimajorAxis interface provided by the
# # AgClassicalSizeShapeSemimajorAxis object. 

# keplarian.SizeShapeType = STKObjects.eSizeShapeSemimajorAxis
# keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).SemiMajorAxis = 7159
# keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).Eccentricity = 0

# ##############################################################################
# ##############################################################################

# keplarian.Orientation.Inclination = 86.4
# keplarian.Orientation.ArgOfPerigee = 0

# # For the RAAN, much as in the case of the semi-major axis and eccentricity,
# # you must first specify the AscNodeType, then provide the value for the
# # AscNode through the approriate interface.
# keplarian.Orientation.AscNodeType = STKObjects.eAscNodeRAAN
# keplarian.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = 45

# ##############################################################################
# ##############################################################################

# keplarian.LocationType = STKObjects.eLocationTrueAnomaly
# keplarian.Location.QueryInterface(STKObjects.IAgClassicalLocationTrueAnomaly).Value = 45

# ##############################################################################
# ##############################################################################

# twoBodyPropagator.InitialState.Representation.Assign(keplarian)
# twoBodyPropagator.Propagate()

# ##############################################################################
# ##############################################################################

# #Remove the test Satellite
# satellite.Unload()

#Insert the constellation of Satellites
numOrbitPlanes = 4
numSatsPerPlane = 8

for orbitPlaneNum, RAAN in enumerate(range(0,180,180//numOrbitPlanes),1): #RAAN in degrees

    for satNum, trueAnomaly in enumerate(range(0,360,360//numSatsPerPlane), 1): #trueAnomaly in degrees
        
        #Insert satellite
        satellite = scenario.Children.New(STKObjects.eSatellite, f"Sat{orbitPlaneNum}{satNum}")
        satellite2 = satellite.QueryInterface(STKObjects.IAgSatellite)
        
        #Change some basic display attributes
        satelliteBasicGfxAttributes =  satellite2.Graphics.Attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)
        satelliteBasicGfxAttributes.Color = 65535 #Yellow
        satelliteBasicGfxAttributes.Line.Width = STKObjects.e2     
        satelliteBasicGfxAttributes.Inherit = False
        satelliteBasicGfxAttributes.QueryInterface(STKObjects.IAgVeGfxAttributesOrbit).IsGroundTrackVisible = False
                
        #Select Propagator
        satellite2.SetPropagatorType(STKObjects.ePropagatorTwoBody)
        
        #Set initial state
        twoBodyPropagator = satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)
        keplarian = twoBodyPropagator.InitialState.Representation.ConvertTo(STKUtil.eOrbitStateClassical).QueryInterface(STKObjects.IAgOrbitStateClassical)
        
        keplarian.SizeShapeType = STKObjects.eSizeShapeSemimajorAxis
        keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).SemiMajorAxis = 7159 #km
        keplarian.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis).Eccentricity = 0

        keplarian.Orientation.Inclination = 86.4 #degrees
        keplarian.Orientation.ArgOfPerigee = 0 #degrees
        keplarian.Orientation.AscNodeType = STKObjects.eAscNodeRAAN
        keplarian.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = RAAN  #degrees
        
        keplarian.LocationType = STKObjects.eLocationTrueAnomaly
        keplarian.Location.QueryInterface(STKObjects.IAgClassicalLocationTrueAnomaly).Value = trueAnomaly + (360//numSatsPerPlane/2)*(orbitPlaneNum%2)  #Stagger true anomalies (degrees) for every other orbital plane       
        
        #Propagate
        satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).InitialState.Representation.Assign(keplarian)
        satellite2.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).Propagate()
    


#############################################################################
#############################################################################

# You can quickly get a collection of all the satellites in the scenario,
# through the Children property of scenario. As you saw before, the Children
# property provides a pointer to the IAgStkObjectCollection interface. This
# interface implements the GetElements, which "Returns a collection of objects
# of specified type." You can then iterate over each of the objects in the
# returned collection (see the GetElements method in line 378).

# Create a new Constellation Object
sensorConstellation = scenario.Children.New(STKObjects.eConstellation, "SensorConstellation")

# Get a pointer to the IAgConstellation interface
sensorConstellation2 = sensorConstellation.QueryInterface(STKObjects.IAgConstellation)

#Loop over all satellites
for satellite in scenario.Children.GetElements(STKObjects.eSatellite):
        
    #Attach sensors to the satellite
    sensor = satellite.Children.New(STKObjects.eSensor,f"Sensor{satellite.InstanceName[3:]}")
    
########## ACTION 3 : Get a pointer to the IAgSensor interface ##########
    sensor2 = sensor.QueryInterface(STKObjects.IAgSensor)
    
    #Adjust Half Cone Angle
    sensor2.CommonTasks.SetPatternSimpleConic(62.5, 2)

    #Adjust the translucency of the sensor projections and the line style
    sensor2.VO.PercentTranslucency = 75
    sensor2.Graphics.LineStyle = STKUtil.eDotted
    
    #Add the sensor to the SensorConstellation
    sensorConstellation2.Objects.Add(sensor.Path)

##############################################################################
##############################################################################

#Create Facility Constellation
facilityConstellation = scenario.Children.New(STKObjects.eConstellation, "FacilityConstellation")
facilityConstellation2 = facilityConstellation.QueryInterface(STKObjects.IAgConstellation)

#Loop over each facility
for facility in scenario.Children.GetElements(STKObjects.eFacility):
    facilityConstellation2.Objects.Add(facility.Path)

#Create chain
chain = scenario.Children.New(STKObjects.eChain, "FacsToSensors")
chain2 = chain.QueryInterface(STKObjects.IAgChain)

#Edit some chain graphics properties
chain2.Graphics.Animation.Color = 65280 #Green
chain2.Graphics.Animation.LineWidth = STKObjects.e3
chain2.Graphics.Animation.IsHighlightVisible = False

#Add objects to chain and compute access
chain2.Objects.Add(facilityConstellation.Path)
chain2.Objects.Add(sensorConstellation.Path)
chain2.ComputeAccess()

# ##############################################################################
# ##############################################################################

facilityAccess = chain.DataProviders.Item('Object Access').QueryInterface(STKObjects.IAgDataPrvInterval).Exec(scenario2.StartTime,scenario2.StopTime)

# ##############################################################################
# ##############################################################################

print("\nObject Access Data Provider, Intervals Count and Items:")
print(facilityAccess.Intervals.Count)
print(facilityAccess.Intervals.Item(0).DataSets.GetRow(0))
print(facilityAccess.Intervals.Item(1).DataSets.GetRow(0))

##############################################################################
##############################################################################

# Count the number of facilities
facilityCount = scenario.Children.GetElements(STKObjects.eFacility).Count

print("\nFacility access data")
for facilityNum in range(facilityCount):
    facilityDataSet = facilityAccess.Intervals.Item(facilityNum).DataSets
    
    el = facilityDataSet.ElementNames

    numRows = facilityDataSet.RowCount
    
    with open(f"Fac{facilityNum+1:02}Access.txt", "w") as dataFile:
        dataFile.write(f"{el[0]},{el[2]},{el[3]},{el[4]}\n")
        
        for row in range(numRows):
            rowData = facilityDataSet.GetRow(row)
            dataFile.write(f"{rowData[0]},{rowData[2]},{rowData[3]},{rowData[4]}\n")
            
    dataFile.close()
    
    # Deletes old information from MaxOutageData.txt
    if facilityNum == 0:
        if os.path.exists("MaxOutageData.txt"):
            open('MaxOutageData.txt', 'w').close()
        
    maxOutage=None    
    with open("MaxOutageData.txt", "a") as outageFile:
       
        #If only one row of data, coverage is continuous
        if numRows == 1:
            outageFile.write(f"Fac{facilityNum+1:02},NA,NA,NA\n")
            print(f"Fac{facilityNum+1:02}: No Outage")
        
        else:
            #Get StartTimes and StopTimes as lists
            startTimes = list(facilityDataSet.GetDataSetByName("Start Time").GetValues())
            stopTimes = list(facilityDataSet.GetDataSetByName("Stop Time").GetValues())
            
            #convert from strings to datetimes, and create np arrays
            startDatetimes = np.array([dt.datetime.strptime(startTime[:-3], "%d %b %Y %H:%M:%S.%f") for startTime in startTimes])
            stopDatetimes = np.array([dt.datetime.strptime(stopTime[:-3], "%d %b %Y %H:%M:%S.%f") for stopTime in stopTimes])
            
            #Compute outage times
            outages = startDatetimes[1:] - stopDatetimes[:-1]
            
            #Locate max outage and associated start and stop time
            maxOutage = np.amax(outages).total_seconds()
            start = stopTimes[np.argmax(outages)]
            stop = startTimes[np.argmax(outages)+1]
            
            #Write out maxoutage data
            outageFile.write(f"Fac{facilityNum+1:02},{maxOutage},{start},{stop}\n")
            print(f"Fac{facilityNum+1:02}: {maxOutage} seconds from {start} until {stop}")
    
    outageFile.close()

    
##############################################################################
##############################################################################

#Get FacTwo object
facTwo = scenario.Children.Item("Fac02")

#Add and configure constraint
facTwoConstraints = facTwo.AccessConstraints
facTwoAzConstraint = facTwoConstraints.AddConstraint(STKObjects.eCstrAzimuthAngle).QueryInterface(STKObjects.IAgAccessCnstrMinMax)

facTwoAzConstraint.EnableMin = True

########## ACTION 1 : Replace ? with the property that will enable the Max property ##########
facTwoAzConstraint.EnableMax = True

facTwoAzConstraint.Min = 45 #degrees
facTwoAzConstraint.Max = 315 #degrees

##############################################################################
##############################################################################

#Compute access
access = facTwo.GetAccess("Satellite/Sat11")
access.ComputeAccess()

##############################################################################
##############################################################################

#Get the access data provider
accessDataPrv = access.DataProviders.Item("Access Data").QueryInterface(STKObjects.IAgDataPrvInterval).Exec(scenario2.StartTime, scenario2.StopTime)

#Get Start Time data and print the first access start time
accessStartTimes = accessDataPrv.DataSets.GetDataSetByName("Start Time").GetValues()
print("\nFac02-Sat11 first access start time:")
print(accessStartTimes[0])

##############################################################################
##############################################################################

#Insert aircraft
aircraft = scenario.Children.New(STKObjects.eAircraft, "TestAircraft")
aircraft2 = aircraft.QueryInterface(STKObjects.IAgAircraft)

##############################################################################
##############################################################################

print("\ndir(aircraft2):")
print(dir(aircraft2))

##############################################################################
##############################################################################

#Change AC attitude to Coordinated turn
attitude = aircraft2.Attitude.QueryInterface(STKObjects.IAgVeRouteAttitudeStandard)
attitude.Basic.SetProfileType(STKObjects.eCoordinatedTurn)

##############################################################################
##############################################################################

#Compute aircraft start time
convertUtil = stkRoot.ConversionUtility
aircraftStartTime = convertUtil.NewDate("UTCG",accessStartTimes[0])
aircraftStartTime = aircraftStartTime.Add("min", 30)
print("\nCalculated aircraft start time:")
print(aircraftStartTime.format("UTCG"))

##############################################################################
##############################################################################

# Load waypoint file
waypoints = np.genfromtxt("FlightPlan.txt", skip_header=1, delimiter=",")
print("\nAircraft Waypoints:")
print(waypoints)

##############################################################################
##############################################################################

#Set propagtor to GreatArc
aircraft2.SetRouteType(STKObjects.ePropagatorGreatArc)
route = aircraft2.Route.QueryInterface(STKObjects.IAgVePropagatorGreatArc)

#Set route start time
startEp = route.EphemerisInterval.GetStartEpoch()
startEp.SetExplicitTime(aircraftStartTime.format("UTCG"))
route.EphemerisInterval.SetStartEpoch(startEp)

#Set the calculation method
route.Method = STKObjects.eDetermineTimeAccFromVel

#Set the altitude reference to MSL
route.SetAltitudeRefType(STKObjects.eWayPtAltRefMSL)

##############################################################################
##############################################################################

#Set unit prefs
stkRoot.UnitPreferences.SetCurrentUnit("DistanceUnit","nm")
stkRoot.UnitPreferences.SetCurrentUnit("TimeUnit","hr")

#Add aircraft waypoints to route
for waypoint in waypoints:
    newWaypoint = route.Waypoints.Add()
    newWaypoint.Latitude = waypoint[0] #degree
    newWaypoint.Longitude = waypoint[1] #degree
    newWaypoint.Altitude = convertUtil.ConvertQuantity("DistanceUnit","ft","nm", waypoint[2]) #ft->nm
    newWaypoint.Speed = waypoint[3] #knots
    newWaypoint.TurnRadius = 1.8 #nautical Miles

#Propagate and reset unit prefs
route.Propagate()
stkRoot.UnitPreferences.ResetUnits()

##############################################################################
##############################################################################

# Set some graphics properties of the aircraft
aircraftBasicGfxAttributes = aircraft2.Graphics.Attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)
aircraftBasicGfxAttributes.Color = 16711935 #Magenta
aircraftBasicGfxAttributes.Line.Width = STKObjects.e3

#Switch to C-130 Model
modelFile = aircraft2.VO.Model.ModelData.QueryInterface(STKObjects.IAgVOModelFile)
modelFile.Filename = os.path.abspath(uiApp.Path[:-3] + "STKData\VO\Models\Air\c-130_hercules.mdl")

##############################################################################
##############################################################################

#Add aircraft constraint
aircraftConstraints = aircraft.AccessConstraints

elConstraint = aircraftConstraints.AddConstraint(STKObjects.eCstrElevationAngle).QueryInterface(STKObjects.IAgAccessCnstrMinMax)
elConstraint.EnableMin = True
elConstraint.Min = 10

##############################################################################
##############################################################################

#Insert and configure the degraded sensor constellation
degradeSensorConstellation = sensorConstellation.CopyObject("DegradedSensorConstellation")
degradeSensorConstellation2 = degradeSensorConstellation.QueryInterface(STKObjects.IAgConstellation)
degradeSensorConstellation2.Objects.RemoveName("Satellite/Sat11/Sensor/Sensor11")

##############################################################################
##############################################################################

#Insert New Chain
aircraftChain = scenario.Children.New(STKObjects.eChain, "AcftToSensors")
aircraftChain2 = aircraftChain.QueryInterface(STKObjects.IAgChain)

#Configure chain graphics
aircraftChain2.Graphics.Animation.Color = 65280 #Green
aircraftChain2.Graphics.Animation.LineWidth = STKObjects.e3
aircraftChain2.Graphics.Animation.IsHighlightVisible = False

#Add objects to chain
aircraftChain2.Objects.Add(aircraft.Path)
aircraftChain2.Objects.Add(degradeSensorConstellation.Path)
aircraftChain2.ComputeAccess()

##############################################################################
##############################################################################

########## ACTION IS REQUIRED IN THIS BLOCK ##########
########## 1 ACTION REQUIRED ##########

########## ACTION 1 : Replace ? with the scenario start time ##########
aircraftAccess = aircraftChain.DataProviders.Item("Complete Access").QueryInterface(STKObjects.IAgDataPrvInterval).Exec(scenario2.StartTime,scenario2.StopTime)


el = aircraftAccess.DataSets.ElementNames
numRows = aircraftAccess.DataSets.RowCount

print("\nAircraft chain access data")
with open("AircraftAccess.txt", "w") as dataFile:
    dataFile.write(f"{el[0]},{el[1]},{el[2]},{el[3]}\n")
    print(f"{el[0]},{el[1]},{el[2]},{el[3]}")
    
    for row in range(numRows):
        rowData = aircraftAccess.DataSets.GetRow(row)
        dataFile.write(f"{rowData[0]},{rowData[1]},{rowData[2]},{rowData[3]}\n")
        print(f"{rowData[0]},{rowData[1]},{rowData[2]},{rowData[3]}")
        
if numRows == 1:
    print(f"No Outage")

else:
    #Get StartTimes and StopTimes as lists
    startTimes = list(aircraftAccess.DataSets.GetDataSetByName("Start Time").GetValues())
    stopTimes = list(aircraftAccess.DataSets.GetDataSetByName("Stop Time").GetValues())
    
    #convert from strings to datetimes, and create np arrays
    startDatetimes = np.array([dt.datetime.strptime(startTime[:-3], "%d %b %Y %H:%M:%S.%f") for startTime in startTimes])
    stopDatetimes = np.array([dt.datetime.strptime(stopTime[:-3], "%d %b %Y %H:%M:%S.%f") for stopTime in stopTimes])
    
    #Compute outage times
    outages = startDatetimes[1:] - stopDatetimes[:-1]
    
    #Locate max outage and associated start and stop time
    maxOutage = np.amax(outages).total_seconds()
    start = stopTimes[np.argmax(outages)]
    stop = startTimes[np.argmax(outages)+1]
    
    #Write out maxoutage data
    print(f"\nAC Max Outage: {maxOutage} seconds from {start} until {stop}")
    
##############################################################################
##############################################################################

# Get the aircraft LLA State Data Provider
aircraftLLA = aircraft.DataProviders.Item("LLA State").QueryInterface(STKObjects.IAgDataProviderGroup)

##############################################################################
##############################################################################

#Specify the Fixed Group of the data provider
aircraftLLAFixed = aircraftLLA.Group.Item("Fixed").QueryInterface(STKObjects.IAgDataPrvTimeVar).Exec(scenario2.StartTime, scenario2.StopTime, 600)

##############################################################################
##############################################################################

#Set unit prefs
stkRoot.UnitPreferences.SetCurrentUnit("DistanceUnit","ft")

#Extract desired aircraft LLA data
el = aircraftLLAFixed.DataSets.ElementNames
aircraftLLAFixedRes = np.array(aircraftLLAFixed.DataSets.ToArray())

print("\nAircraft LLA State (Fixed) data:")
print(f"{el[0]:30} {el[1]:20} {el[2]:28} {el[11]:15}")
for lla in aircraftLLAFixedRes:
    print(f"{lla[0]:30} {lla[1]:20} {lla[2]:20} {round(float(lla[11])):15}")

#Reset unit prefs
stkRoot.UnitPreferences.ResetUnits()

##############################################################################
##############################################################################

#Get aircraft All Postion data provider and print the associated data
facTwoPosData = facTwo.DataProviders.Item("All Position").QueryInterface(STKObjects.IAgDataPrvFixed).Exec()
els = facTwoPosData.DataSets.ElementNames
data = facTwoPosData.DataSets.ToArray()[0]

print("\nFac02 Position data")
for idx, el in enumerate(els):
    print(f"{el}: {data[idx]}")

##############################################################################
##############################################################################

#End