import sys
import os.path
import locale
import OrcFxAPIConfig
import ctypes
from ctypes.wintypes import BOOL, HBITMAP, MAX_PATH

_interfaceCapabilities = 1
_isPy3k = sys.version_info[0] >= 3
_lib = OrcFxAPIConfig.lib()


def _redirectOutputToOrcaFlex():
    class outputStream(object):

        def write(self, str):
            status = ctypes.c_int()
            str = _PrepareString(str)
            _fn('C_ExternalFunctionPrint')(_pchar(str), ctypes.byref(status))
            _CheckStatus(status)
    sys.stderr = sys.stdout = outputStream()


def _redirectOutputToNull():
    class outputStream(object):

        def write(self, str):
            pass
    sys.stderr = sys.stdout = outputStream()

_discard, _Encoding = locale.getdefaultlocale()
del _discard

_UseUnicodeOrcFxAPI = hasattr(_lib, 'C_GetLastErrorStringW')
if _UseUnicodeOrcFxAPI:
    _char = ctypes.c_wchar
    _pchar = ctypes.c_wchar_p
else:
    _char = ctypes.c_char
    _pchar = ctypes.c_char_p

Handle = ctypes.c_void_p


def _PreparePythonString(value):
    return str(value)  # str is ASCII in Python 2, and Unicode in Python 3


def _PreparePy3kString(value):
    if isinstance(value, str):
        return value.encode(_Encoding)
    elif isinstance(value, bytes):
        return value
    else:
        return str(value).encode(_Encoding)


def _PreparePy2kStringForUnicodeAPI(value):
    if isinstance(value, str):
        return unicode(value, _Encoding, 'strict')
    else:
        return unicode(value)

if _isPy3k and not _UseUnicodeOrcFxAPI:
    _PrepareString = _PreparePy3kString
elif not _isPy3k and _UseUnicodeOrcFxAPI:
    _PrepareString = _PreparePy2kStringForUnicodeAPI
else:
    _PrepareString = _PreparePythonString


def _DecodeString(value):
    if _isPy3k and isinstance(value, bytes):
        return value.decode(_Encoding)  # byte data to string
    else:
        return value


def _fn(ansiFunctionName):
    return _lib[ansiFunctionName + 'W' if _UseUnicodeOrcFxAPI else ansiFunctionName]

_fn('C_GetDataString').restype = ctypes.c_int
_fn('C_GetLastErrorString').restype = ctypes.c_int
_fn('C_GetObjectTypeName').restype = ctypes.c_int
_fn('C_GetWarningText').restype = ctypes.c_int
_lib.OrcinaDefaultReal.restype = ctypes.c_double
_lib.OrcinaInfinity.restype = ctypes.c_double
_lib.OrcinaUndefinedReal.restype = ctypes.c_double
_lib.OrcinaNullReal.restype = ctypes.c_double
_lib.C_GetNumOfSamples.restype = ctypes.c_int
_lib.C_GetNumOfSamples.restype = ctypes.c_int
_lib.C_GetNumOfWarnings.restype = ctypes.c_int
_lib.C_GetSimulationDrawTime.restype = ctypes.c_double
_lib.C_GetSimulationTimeToGo.restype = ctypes.c_double

if hasattr(_lib, 'C_CalculateExtremeStatisticsExcessesOverThreshold'):
    _lib.C_CalculateExtremeStatisticsExcessesOverThreshold.restype = ctypes.c_int
if hasattr(_lib, 'C_GetModelThreadCount'):
    _lib.C_GetModelThreadCount.restype = ctypes.c_int

_DataCountAvailable = hasattr(_lib, 'C_GetDataRowCount')
_GetWaveComponents2Available = hasattr(_lib, 'C_GetWaveComponents2')
_CreateModel2Available = hasattr(_lib, 'C_CreateModel2')
_BeginEndDataChangeAvailable = hasattr(_lib, 'C_BeginDataChange')
_ReportActionProgressAvailable = _UseUnicodeOrcFxAPI and hasattr(_lib, 'C_ReportActionProgressW')


def _checkVersion(requiredVersion):
    s = (_char * 16)()
    s.value = _PrepareString(requiredVersion)
    OK, status = ctypes.c_int(), ctypes.c_int()
    _fn('C_GetDLLVersion')(
        ctypes.byref(s),
        ctypes.byref((_char * 16)()),
        ctypes.byref(OK),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return OK.value == 1


def SetUseNumpyArray(UseNumpyArray):
    global array
    global NumpyAvailable
    global numpy

    if UseNumpyArray:
        try:
            import numpy
            import numpy.ctypeslib
            NumpyAvailable = True
        except:
            NumpyAvailable = False
    else:
        NumpyAvailable = False

    def array(values):
        result = tuple(values)
        if NumpyAvailable:
            result = numpy.array(result)
        return result

SetUseNumpyArray(True)


def _requiresNumpy():
    if not NumpyAvailable:
        raise MissingRequirementError('CalculateMooringStiffness requires numpy')

# Status constants returned by API functions
stOK = 0
stDLLsVersionError = 1
stCreateModelError = 2
stModelHandleError = 3
stObjectHandleError = 4
stInvalidObjectType = 5
stFileNotFound = 6
stFileReadError = 7
stTimeHistoryError = 8
stNoSuchNodeNum = 9
stInvalidPropertyNum = 10
stInvalidPeriod = 12
stInvalidVarNum = 13
stRangeGraphError = 14
stInvalidObjectExtra = 15
stNotEnoughVars = 16
stInvalidVars = 17
stUnrecognisedVarNum = 18
stInvalidHandle = 19
stUnexpectedError = 20
stInvalidIndex = 21
stNoSuchObject = 23
stNotAVariantArray = 24
stLicensingError = 25
stUnrecognisedVarName = 26
stStaticsFailed = 27
stFileWriteError = 28
stOperationCancelled = 29
stSolveEquationFailed = 30
stInvalidDataName = 31
stInvalidDataType = 32
stInvalidDataAccess = 33
stInvalidVersion = 34
stInvalidStructureSize = 35
stRequiredModulesNotEnabled = 36
stPeriodNotYetStarted = 37
stCouldNotDestroyObject = 38
stInvalidModelState = 39
stSimulationError = 40
stInvalidModule = 41
stInvalidResultType = 42
stInvalidViewParameters = 43
stCannotExtendSimulation = 44
stUnrecognisedObjectTypeName = 45
stUnknownModelState = 46
stFunctionNotAvailable = 47
stStructureSizeTooSmall = 48
stInvalidParameter = 49
stResponseGraphError = 50
stResultsNotAvailableWhenNotIncludedInStatics = 51
stInvalidFileType = 52
stBatchScriptFailed = 53
stInvalidTimeHistoryValues = 54
stResultsNotLogged = 55
stWizardFailed = 56
stDLLInitFailed = 57
stInvalidArclengthRange = 58
stValueNotAvailable = 59
stInvalidValue = 60
stModalAnalysisFailed = 61
stVesselTypeDataImportFailed = 62
stOperationNotAvailable = 63
stFatigueAnalysisFailed = 64
stExtremeStatisticsFailed = 65

# Object type constants
otGeneral = 1
otEnvironment = 3
otVessel = 5
otLine = 6
ot6DBuoy = 7
ot3DBuoy = 8
otWinch = 9
otLink = 10
otShape = 11
otDragChain = 14
otLineType = 15
otClumpType = 16
otWingType = 17
otVesselType = 18
otDragChainType = 19
otFlexJointType = 20
otStiffenerType = 21
otFlexJoint = 41
otAttachedBuoy = 43
otSolidFrictionCoefficients = 47
otRayleighDampingCoefficients = 48
otWakeModel = 49
otPyModel = 50
otLineContact = 51
otCodeChecks = 52
otShear7Data = 53
otVIVAData = 54
otSupportType = 55
otMultibodyGroup = 62

# Object type constants (variable data sources)
otDragCoefficient = 1000
otAxialStiffness = 1001
otBendingStiffness = 1002
otBendingConnectionStiffness = 1003
otWingOrientation = 1004
otKinematicViscosity = 1005
otFluidTemperature = 1006
otCurrentSpeed = 1007
otCurrentDirection = 1008
otExternalFunction = 1009
otHorizontalVariationFactor = 1010
otLoadForce = 1011
otLoadMoment = 1012
otExpansionFactor = 1013
otWinchPayoutRate = 1014
otWinchTension = 1015
otVerticalVariationFactor = 1016
otTorsionalStiffness = 1017
otMinimumBendRadius = 1018
otLiftCoefficient = 1019
otLiftCloseToSeabed = 1020
otDragCloseToSeabed = 1021
otDragAmplificationFactor = 1022
otLineTypeDiameter = 1023
otStressStrainRelationship = 1024
otCoatingOrLining = 1025
otContentsFlowVelocity = 1026
otAddedMassRateOfChangeCloseToSurface = 1027
otAddedMassCloseToSurface = 1028
otContactStiffness = 1029
otSupportsStiffness = 1030

# Variable Data Names
vdnWingAzimuth = 'WingAzimuth'
vdnWingDeclination = 'WingDeclination'
vdnWingGamma = 'WingGamma'
vdnGlobalAppliedForceX = 'GlobalAppliedForceX'
vdnGlobalAppliedForceY = 'GlobalAppliedForceY'
vdnGlobalAppliedForceZ = 'GlobalAppliedForceZ'
vdnGlobalAppliedMomentX = 'GlobalAppliedMomentX'
vdnGlobalAppliedMomentY = 'GlobalAppliedMomentY'
vdnGlobalAppliedMomentZ = 'GlobalAppliedMomentZ'
vdnLocalAppliedForceX = 'LocalAppliedForceX'
vdnLocalAppliedForceY = 'LocalAppliedForceY'
vdnLocalAppliedForceZ = 'LocalAppliedForceZ'
vdnLocalAppliedMomentX = 'LocalAppliedMomentX'
vdnLocalAppliedMomentY = 'LocalAppliedMomentY'
vdnLocalAppliedMomentZ = 'LocalAppliedMomentZ'
vdnRefCurrentSpeed = 'RefCurrentSpeed'
vdnRefCurrentDirection = 'RefCurrentDirection'
vdnWholeSimulationTension = 'WholeSimulationTension'
vdnWholeSimulationPayoutRate = 'WholeSimulationPayoutRate'
vdnXBendStiffness = 'xBendStiffness'
vdnXBendMomentIn = 'xBendMomentIn'
vdnYBendMomentIn = 'yBendMomentIn'
vdnXBendMomentOut = 'xBendMomentOut'
vdnYBendMomentOut = 'yBendMomentOut'
vdnExternallyCalculatedPrimaryMotion = 'ExternallyCalculatedPrimaryMotion'

# Period constants
pnBuildUp = 0
# Stage n has Period number n
pnSpecifiedPeriod = 32001
pnLatestWave = 32002
pnWholeSimulation = 32003
pnStaticState = 32004
pnInstantaneousValue = 32005

# For ObjectExtra.LinePoint
ptEndA = 0
ptEndB = 1
ptTouchdown = 2
ptNodeNum = 3
ptArcLength = 4

# For ObjectExtra.RadialPos
rpInner = 0
rpOuter = 1

# Special integer value equivalent to '~' in OrcaFlex UI
OrcinaDefaultWord = 65500

# DataType constants
dtDouble = 0
dtInteger = 1
dtString = 2
dtVariable = 3

# For EnumerateVars
rtTimeHistory = 0
rtRangeGraph = 1
rtLinkedStatistics = 2
rtFrequencyDomain = 3

# For range graph arclength range modes
armEntireLine = 0
armSpecifiedArclengths = 1
armSpecifiedSections = 2

# For move objects
sbDisplacement = 0
sbPolarDisplacement = 1
sbNewPosition = 2
sbRotation = 3

# For module disable/enabled
moduleDynamics = 0
moduleVIV = 1

# For time history summary
thstSpectralDensity = 0
thstEmpiricalDistribution = 1
thstRainflowHalfCycles = 2

# For view parameter graphics mode
gmWireFrame = 0
gmShaded = 1

# For view parameter shaded fill mode
fmSolid = 0
fmMesh = 1

# For saving external program files
eftShear7dat = 0
eftShear7mds = 1
eftShear7out = 2
eftShear7plt = 3
eftShear7anm = 4
eftShear7dmg = 5
eftShear7fat = 6
eftShear7out1 = 7
eftShear7out2 = 8
eftShear7allOutput = 9
eftShear7str = 10
eftVIVAInput = 11
eftVIVAOutput = 12
eftVIVAModes = 13

# For saving spreadsheets
sptSummaryResults = 0
sptFullResults = 1
sptWaveSearch = 2
sptVesselDisplacementRAOs = 3
sptVesselSpectralResponse = 4
sptLineClashingReport = 5
sptDetailedProperties = 6
sptLineTypesPropertiesReport = 7
sptCodeChecksProperties = 8

# For view parameter file format
bffWindowsBitmap = 0
bffPNG = 1
bffGIF = 2
bffJPEG = 3
bffPDF = 4

# For extreme statistics
evdRayleigh = 0
evdWeibull = 1
evdGPD = 2

exUpperTail = 0
exLowerTail = 1

# For post calculation actions
atInProcPython = 0
atCmdScript = 1

# For modal analysis
mtNotAvailable = -1
mtTransverse = 0
mtMostlyTransverse = 1
mtInline = 2
mtMostlyInline = 3
mtAxial = 4
mtMostlyAxial = 5
mtMixed = 6
mtRotational = 7
mtMostlyRotational = 8

# Model property IDs
propIsTimeDomainDynamics = 0
propIsFrequencyDomainDynamics = 1
propIsDeterministicFrequencyDomainDynamics = 2

# For vessel type data import
vdtDisplacementRAOs = 0
vdtLoadRAOs = 1
vdtNewmanQTFs = 2
vdtFullQTFs = 3
vdtStiffnessAddedMassDamping = 4
vdtMassInertia = 5
vdtOtherDamping = 6
vdtSeaStateRAOs = 7

iftGeneric = 0
iftAQWA = 1
iftWAMIT = 2

# Default autosave interval
DefaultAutoSaveIntervalMinutes = 60  # same as OrcaFlex default


def _CheckStatus(status):
    if status.value != stOK:
        raise DLLError(status.value)


def _GetDataString(handle, name, index):
    func = _fn('C_GetDataString')
    name = _PrepareString(name)
    status = ctypes.c_int()
    length = func(
        handle,
        name,
        index,
        None,
        ctypes.byref(status)
    )
    if status.value == stValueNotAvailable:
        return None
    _CheckStatus(status)
    result = (_char * length)()
    func(
        handle,
        name,
        index,
        ctypes.byref(result),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return _DecodeString(result.value)


def _SaveSpreadsheet(handle, spreadSheetType, filename, parameters):
    if parameters is not None:
        parameters = ctypes.byref(parameters)
    filename = _PrepareString(filename)
    status = ctypes.c_int()
    _fn('C_SaveSpreadsheet')(
        handle,
        ctypes.c_int(spreadSheetType),
        parameters,
        _pchar(filename),
        ctypes.byref(status)
    )
    _CheckStatus(status)


def _ProgressCallbackCancel(progressHandlerReturnValue):
    if progressHandlerReturnValue is None:
        progressHandlerReturnValue = False
    return BOOL(progressHandlerReturnValue)


class _DictLookup(object):

    def __init__(self, dict):
        for (k, v) in dict.items():
            setattr(self, k, v)

    def __repr__(self):
        result = ''
        for name, value in vars(self).items():
            result += '%r: %r, ' % (name, value)
        return '<' + result[0:len(result) - 2] + '>'

    def __str__(self):
        result = ''
        for name, value in vars(self).items():
            result += '%s: %s, ' % (name, value)
        return '<' + result[0:len(result) - 2] + '>'


def objectFromDict(dict):
    return _DictLookup(dict)


class DLLError(Exception):

    def __init__(self, status):
        if isinstance(status, str):
            self.msg = status
        else:
            self.status = status
            func = _fn('C_GetLastErrorString')
            length = func(None)
            ErrorString = (_char * length)()
            func(ctypes.byref(ErrorString))
            self.msg = '\nError code: %d\n%s' % (status, ErrorString.value)

    def __str__(self):
        return self.__repr__()

    def __repr__(self):
        return self.msg


class MissingRequirementError(Exception):
    pass


def Enum(*names):

    class EnumClass(object):
        __slots__ = names

        def __iter__(self):
            return iter(constants)

        def __len__(self):
            return len(constants)

        def __getitem__(self, i):
            return constants[i]

        def __repr__(self):
            return 'Enum' + str(names)

        def __str__(self):
            return 'enum ' + str(constants)

        def __contains__(self, item):
            return item in names

    enumType = EnumClass()

    class EnumValue(object):
        __slots__ = ('__value')

        def __init__(self, value):
            self.__value = value
        Value = property(
            lambda self: self.__value
        )
        EnumType = property(
            lambda self: enumType
        )

        def __hash__(self):
            return hash(self.__value)

        def CheckComparable(self, other):
            if self.EnumType is not other.EnumType:
                raise TypeError('Only values from the same enum are comparable')

        def __eq__(self, other):
            self.CheckComparable(other)
            return self.__value == other.__value

        def __ne__(self, other):
            self.CheckComparable(other)
            return self.__value != other.__value

        def __lt__(self, other):
            self.CheckComparable(other)
            return self.__value < other.__value

        def __gt__(self, other):
            self.CheckComparable(other)
            return self.__value > other.__value

        def __le__(self, other):
            self.CheckComparable(other)
            return self.__value <= other.__value

        def __ge__(self, other):
            self.CheckComparable(other)
            return self.__value >= other.__value

        def __repr__(self):
            return str(names[self.__value])

    maximum = len(names) - 1
    constants = [None] * len(names)
    for i, each in enumerate(names):
        val = EnumValue(i)
        setattr(EnumClass, each, val)
        constants[i] = val
    constants = tuple(constants)
    return enumType

ModelState = Enum(
    'Reset',
    'CalculatingStatics',
    'InStaticState',
    'RunningSimulation',
    'SimulationStopped',
    'SimulationStoppedUnstable'
)

FileType = Enum(
    'DataFile',
    'StaticStateSimulationFile',
    'DynamicSimulationFile'
)


class IndexedDataItem(object):

    def __init__(self, dataName, obj):
        self.dataName = dataName
        self.obj = obj

    def Assign(self, value):
        with self.obj.dataChange():
            count = len(value)
            if len(self) != count:
                self.obj.SetDataRowCount(self.dataName, count)
            for i in range(count):
                self[i] = value[i]

    def wrappedIndex(self, index):
        if not isinstance(index, int):
            raise TypeError('index must be either slice or int instances')
        length = len(self)
        if index < 0:
            index += length
        if index < 0 or index >= length:
            raise IndexError(index)
        return index

    def sliceCount(self, slice):
        start, stop, stride = slice.indices(len(self))
        # (a > b) - (a < b) is equivalent to cmp(a, b), see https://docs.python.org/3.0/whatsnew/3.0.html#ordering-comparisons
        return (stop - start + stride + (0 > stride) - (0 < stride)) // stride

    def sliceIndices(self, slice):
        start, stop, stride = slice.indices(len(self))
        if stride > 0:
            moreItems = lambda index, stop: index < stop
        else:
            moreItems = lambda index, stop: index > stop
        index = start
        while moreItems(index, stop):
            yield index
            index += stride

    def __getitem__(self, index):
        if isinstance(index, slice):
            return [self.obj.GetData(self.dataName, i) for i in self.sliceIndices(index)]
        else:
            return self.obj.GetData(self.dataName, self.wrappedIndex(index))

    def __setitem__(self, index, value):
        if isinstance(index, slice):
            if self.sliceCount(index) != len(value):
                raise ValueError(
                    'attempt to assign sequence of size %d to slice of size %d' %
                    (len(value), self.sliceCount(index)))
            for i, item in zip(self.sliceIndices(index), value):
                self.obj.SetData(self.dataName, i, item)
        else:
            self.obj.SetData(self.dataName, self.wrappedIndex(index), value)

    def __len__(self):
        return self.obj.GetDataRowCount(self.dataName)

    def __repr__(self):
        return repr(tuple(self))

    def DeleteRow(self, index):
        self.obj.DeleteDataRow(self.dataName, index)

    def InsertRow(self, index):
        self.obj.InsertDataRow(self.dataName, index)

    @property
    def rowCount(self):
        return self.obj.GetDataRowCount(self.dataName)

    @rowCount.setter
    def rowCount(self, value):
        self.obj.SetDataRowCount(self.dataName, value)


class PackedStructure(ctypes.Structure):

    _pack_ = 1

    def __init__(self, **kwargs):
        ctypes.Structure.__init__(self, **kwargs)
        if hasattr(self, 'Size'):
            self.Size = ctypes.sizeof(self)

    def asObject(self):
        dict = {}
        for field, type in self._fields_:
            if field != 'Size':
                dict[field] = getattr(self, field)
        return objectFromDict(dict)


class PackedStructureWithObjectHandles(PackedStructure):

    @staticmethod
    def _objectFromHandle(handle):
        objectHandle = Handle(handle)
        if objectHandle == 0:
            return None
        modelHandle, status = Handle(), ctypes.c_int()
        _lib.C_GetModelHandle(
            objectHandle,
            ctypes.byref(modelHandle),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return Model(handle=modelHandle).createOrcaFlexObject(objectHandle)

    @staticmethod
    def _handleFromObject(value):
        if value is None:
            return None
        else:
            return value.handle


class CreateModelParams(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('ThreadCount', ctypes.c_int)
    ]


class ObjectInfo(PackedStructure):

    _fields_ = [
        ('ObjectHandle', Handle),
        ('ObjectType', ctypes.c_int),
        ('ObjectName', _char * 50)
    ]


class Period(PackedStructure):

    _fields_ = [
        ('PeriodNum', ctypes.c_int),
        ('Unused', ctypes.c_int),
        ('FromTime', ctypes.c_double),
        ('ToTime', ctypes.c_double)
    ]

    def __init__(self, PeriodNum, FromTime=0.0, ToTime=0.0):
        PackedStructure.__init__(self)
        self.PeriodNum = PeriodNum
        self.FromTime = FromTime
        self.ToTime = ToTime


def SpecifiedPeriod(FromTime=0.0, ToTime=0.0):
    return Period(pnSpecifiedPeriod, FromTime, ToTime)


class ObjectExtra(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('EnvironmentPos', ctypes.c_double * 3),
        ('LinePoint', ctypes.c_int),
        ('NodeNum', ctypes.c_int),
        ('ArcLength', ctypes.c_double),
        ('RadialPos', ctypes.c_int),
        ('Theta', ctypes.c_double),
        ('WingName', _pchar),
        ('ClearanceLineName', _pchar),
        ('WinchConnectionPoint', ctypes.c_int),
        ('RigidBodyPos', ctypes.c_double * 3),
        ('ExternalResultText', _pchar),
        ('DisturbanceVesselName', _pchar),
        ('SupportIndex', ctypes.c_int),
        ('SupportedLineName', _pchar)
    ]


def OrcinaDefaultReal():
    return _lib.OrcinaDefaultReal()


def OrcinaInfinity():
    return _lib.OrcinaInfinity()


def OrcinaUndefinedReal():
    return _lib.OrcinaUndefinedReal()


def OrcinaNullReal():
    return _lib.OrcinaNullReal()


def Vector(*args):
    if len(args) == 1:
        args = args[0]  # assume parameter is iterable
    if len(args) != 3:
        raise TypeError('Vector must have length of 3')
    return list(args)


def Vector2(*args):
    if len(args) == 1:
        args = args[0]  # assume parameter is iterable
    if len(args) != 2:
        raise TypeError('Vector2 must have length of 2')
    return list(args)


def oeEnvironment(*args):
    result = ObjectExtra()
    result.EnvironmentPos[:] = Vector(*args)
    return result


def oeBuoy(*args):
    result = ObjectExtra()
    result.RigidBodyPos[:] = Vector(*args)
    return result


def oeWing(WingName):
    result = ObjectExtra()
    result.WingName = WingName
    return result


def oeVessel(*args):
    result = ObjectExtra()
    result.RigidBodyPos[:] = Vector(*args)
    return result


def oeSupport(supportIndex, supportedLineName=None):
    result = ObjectExtra()
    result.SupportIndex = supportIndex
    result.SupportedLineName = supportedLineName
    return result


def oeWinch(WinchConnectionPoint):
    result = ObjectExtra()
    result.WinchConnectionPoint = WinchConnectionPoint
    return result


def oeLine(
        LinePoint=ptArcLength,
        NodeNum=0,
        ArcLength=0.0,
        RadialPos=rpInner,
        Theta=0.0,
        ClearanceLineName=None,
        ExternalResultText=None):
    result = ObjectExtra()
    result.LinePoint = LinePoint
    result.NodeNum = NodeNum
    result.ArcLength = ArcLength
    result.RadialPos = RadialPos
    result.Theta = Theta
    result.ClearanceLineName = ClearanceLineName
    result.ExternalResultText = ExternalResultText
    return result


def oeNodeNum(NodeNum):
    return oeLine(LinePoint=ptNodeNum, NodeNum=NodeNum)


def oeArcLength(ArcLength):
    return oeLine(LinePoint=ptArcLength, ArcLength=ArcLength)

oeEndA = oeLine(LinePoint=ptEndA)
oeEndB = oeLine(LinePoint=ptEndB)
oeTouchdown = oeLine(LinePoint=ptTouchdown)


class RangeGraphCurveNames(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Min', _char * 30),
        ('Max', _char * 30),
        ('Mean', _char * 30),
        ('StdDev', _char * 30),
        ('Upper', _char * 30),
        ('Lower', _char * 30)
    ]


class ArclengthRange(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Mode', ctypes.c_int),
        ('FromArclength', ctypes.c_double),
        ('ToArclength', ctypes.c_double),
        ('FromSection', ctypes.c_int),
        ('ToSection', ctypes.c_int)
    ]


def arEntireLine():
    result = ArclengthRange()
    result.Mode = armEntireLine
    return result


def arSpecifiedArclengths(FromArclength, ToArclength):
    result = ArclengthRange()
    result.Mode = armSpecifiedArclengths
    result.FromArclength = FromArclength
    result.ToArclength = ToArclength
    return result


def arSpecifiedSections(FromSection, ToSection):
    result = ArclengthRange()
    result.Mode = armSpecifiedSections
    result.FromSection = FromSection
    result.ToSection = ToSection
    return result


class MoveObjectPoint(PackedStructureWithObjectHandles):

    _fields_ = [
        ('ObjectHandle', Handle),
        ('PointIndex', ctypes.c_int)
    ]

    def __init__(self, object, pointIndex):
        PackedStructureWithObjectHandles.__init__(self)
        self.ObjectHandle = object.handle
        self.PointIndex = pointIndex

    @property
    def Object(self):
        return self._objectFromHandle(self.ObjectHandle)

    @Object.setter
    def Object(self, value):
        self.ObjectHandle = self._handleFromObject(value)


class MoveObjectSpecification(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('MoveSpecifiedBy', ctypes.c_int),
        ('Displacement', ctypes.c_double * 3),
        ('PolarDisplacementDirection', ctypes.c_double),
        ('PolarDisplacementDistance', ctypes.c_double),
        ('NewPositionReferencePoint', MoveObjectPoint),
        ('NewPosition', ctypes.c_double * 3),
        ('RotationAngle', ctypes.c_double),
        ('RotationCentre', ctypes.c_double * 2)
    ]


class ModesFilesParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('FirstMode', ctypes.c_int),
        ('LastMode', ctypes.c_int)
    ]

Shear7MdsFileParameters = ModesFilesParameters
VIVAModesFilesParameters = ModesFilesParameters


class TimeSeriesStats(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Mean', ctypes.c_double),
        ('StdDev', ctypes.c_double),
        ('m0', ctypes.c_double),
        ('m2', ctypes.c_double),
        ('m4', ctypes.c_double),
        ('Tz', ctypes.c_double),
        ('Tc', ctypes.c_double),
        ('Bandwidth', ctypes.c_double)
    ]


class FrequencyDomainResults(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('StaticValue', ctypes.c_double),
        ('StdDev', ctypes.c_double),
        ('Amplitude', ctypes.c_double),
        ('PhaseLag', ctypes.c_double),
        ('m0', ctypes.c_double),
        ('m1', ctypes.c_double),
        ('m2', ctypes.c_double),
        ('m3', ctypes.c_double),
        ('m4', ctypes.c_double),
        ('Tz', ctypes.c_double),
        ('Tc', ctypes.c_double),
        ('Bandwidth', ctypes.c_double)
    ]


class UseCalculatedPositionsForStaticsParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('SetLinesToUserSpecifiedStartingShape', BOOL)
    ]


class ViewParameters(PackedStructureWithObjectHandles):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('ViewSize', ctypes.c_double),
        ('ViewAzimuth', ctypes.c_double),
        ('ViewElevation', ctypes.c_double),
        ('ViewCentre', ctypes.c_double * 3),
        ('Height', ctypes.c_int),
        ('Width', ctypes.c_int),
        ('BackgroundColour', ctypes.c_int),
        ('DrawViewAxes', BOOL),
        ('DrawScaleBar', BOOL),
        ('DrawGlobalAxes', BOOL),
        ('DrawEnvironmentAxes', BOOL),
        ('DrawLocalAxes', BOOL),
        ('DrawOutOfBalanceForces', BOOL),
        ('DrawNodeAxes', BOOL),
        ('GraphicsMode', ctypes.c_int),
        ('FileFormat', ctypes.c_int)
    ]
    if _checkVersion('9.5a28'):
        _fields_.extend([
            ('ViewGamma', ctypes.c_double),
            ('RelativeToObjectHandle', Handle)
        ])
    if _checkVersion('9.7a14'):
        _fields_.extend([
            ('DisturbanceVesselHandle', Handle),
            ('DisturbancePosition', ctypes.c_double * 2)
        ])
    if _checkVersion('9.8a6'):
        _fields_.extend([
            ('ShadedFillMode', ctypes.c_int)
        ])

    @property
    def RelativeToObject(self):
        return self._objectFromHandle(self.RelativeToObjectHandle)

    @RelativeToObject.setter
    def RelativeToObject(self, value):
        self.RelativeToObjectHandle = self._handleFromObject(value)

    @property
    def DisturbanceVessel(self):
        return self._objectFromHandle(self.DisturbanceVesselHandle)

    @DisturbanceVessel.setter
    def DisturbanceVessel(self, value):
        self.DisturbanceVesselHandle = self._handleFromObject(value)


class AVIFileParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Codec', ctypes.c_uint32),
        ('Interval', ctypes.c_double)
    ]

    def __init__(self, codec, interval):
        PackedStructure.__init__(self)
        self.Codec = ctypes.windll.winmm.mmioStringToFOURCCA(
            str(codec).encode(_Encoding),
            0x0010  # MMIO_TOUPPER = 0x0010
        )
        self.Interval = float(interval)


class VarInfo(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('VarID', ctypes.c_int),
        ('VarName', _pchar),
        ('VarUnits', _pchar),
        ('FullName', _pchar),
        ('ObjectHandle', Handle)
    ]


class SimulationTimeStatus(PackedStructure):

    _fields_ = [
        ('StartTime', ctypes.c_double),
        ('StopTime', ctypes.c_double),
        ('CurrentTime', ctypes.c_double)
    ]


class TimeSteps(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('InnerTimeStep', ctypes.c_double),
        ('OuterTimeStep', ctypes.c_double)
    ]


class TimeHistorySummarySpecification(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('SpectralDensityFundamentalFrequency', ctypes.c_double)
    ]


class RunSimulationParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('EnableAutoSave', BOOL),
        ('AutoSaveIntervalMinutes', ctypes.c_int),
        ('AutoSaveFileName', _pchar)
    ]

    def __init__(self, enableAutoSave, autoSaveIntervalMinutes, autoSaveFileName):
        PackedStructure.__init__(self)
        self.EnableAutoSave = enableAutoSave
        self.AutoSaveIntervalMinutes = autoSaveIntervalMinutes
        self.AutoSaveFileName = _pchar(_PrepareString(autoSaveFileName))


class StatisticsQuery(PackedStructure):

    _fields_ = [
        ('StdDev', ctypes.c_double),
        ('Mean', ctypes.c_double),
        ('TimeOfMax', ctypes.c_double),
        ('ValueAtMax', ctypes.c_double),
        ('LinkedValueAtMax', ctypes.c_double),
        ('TimeOfMin', ctypes.c_double),
        ('ValueAtMin', ctypes.c_double),
        ('LinkedValueAtMin', ctypes.c_double)
    ]


class WaveComponent(PackedStructure):

    if _GetWaveComponents2Available:
        _fields_ = [
            ('WaveTrainIndex', ctypes.c_int),
            ('Frequency', ctypes.c_double),
            ('FrequencyLowerBound', ctypes.c_double),
            ('FrequencyUpperBound', ctypes.c_double),
            ('Amplitude', ctypes.c_double),
            ('PhaseLagWrtWaveTrainTime', ctypes.c_double),
            ('PhaseLagWrtSimulationTime', ctypes.c_double),
            ('WaveNumber', ctypes.c_double),
            ('Direction', ctypes.c_double)
        ]
    else:
        _fields_ = [
            ('WaveTrainIndex', ctypes.c_int),
            ('Frequency', ctypes.c_double),
            ('Amplitude', ctypes.c_double),
            ('PhaseLagWrtWaveTrainTime', ctypes.c_double),
            ('PhaseLagWrtSimulationTime', ctypes.c_double),
            ('WaveNumber', ctypes.c_double),
            ('Direction', ctypes.c_double)
        ]


class WindComponent(PackedStructure):

    _fields_ = [
        ('Frequency', ctypes.c_double),
        ('FrequencyLowerBound', ctypes.c_double),
        ('FrequencyUpperBound', ctypes.c_double),
        ('Amplitude', ctypes.c_double),
        ('PhaseLagWrtWindTime', ctypes.c_double),
        ('PhaseLagWrtSimulationTime', ctypes.c_double),
    ]


class _GraphCurve(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('_X', ctypes.c_void_p),
        ('_Y', ctypes.c_void_p)
    ]

    def __init__(self, count):
        PackedStructure.__init__(self)
        self.X = (ctypes.c_double * count)()
        self.Y = (ctypes.c_double * count)()
        self._X = ctypes.cast(ctypes.byref(self.X), ctypes.c_void_p)
        self._Y = ctypes.cast(ctypes.byref(self.Y), ctypes.c_void_p)


class LineClashingReportParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Period', Period)
    ]


class VesselTypeDataImportSpecification(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('ImportFileType', ctypes.c_int),
        ('BodyCount', ctypes.c_int),
        ('MultibodyGroupName', _pchar),
        ('ImportDisplacementRAOs', BOOL),
        ('ImportLoadRAOs', BOOL),
        ('ImportNewmanQTFs', BOOL),
        ('ImportFullQTFs', BOOL),
        ('ImportStiffnessAddedMassDamping', BOOL),
        ('ImportMassInertia', BOOL),
        ('ImportSeaStateRAOs', BOOL),
        ('ImportOtherDamping', BOOL),
        ('WamitLoadRAOMethod', ctypes.c_int),
        ('WamitDiagonalQTFMethod', ctypes.c_int),
        ('WamitFullQTFMethod', ctypes.c_int)
    ]

    def __init__(self, fileType, bodyCount, multibodyGroup, requestedData, wamitCalculationMethods):

        wamitLoadRAOMethods = Enum(
            'wlmDefault',
            'wlmHaskind',
            'wlmDiffraction'
        )
        wamitDiagonalQTFMethods = Enum(
            'wqmDefault',
            'wqmPressureIntegration',
            'wqmControlSurface',
            'wqmMomentumConservation'
        )
        wamitFullQTFMethods = Enum(
            'wfmDefault',
            'wfmDirect',
            'wfmIndirect'
        )

        PackedStructure.__init__(self)
        self.ImportFileType = fileType
        self.BodyCount = bodyCount
        self.MultibodyGroupName = multibodyGroup
        self.ImportDisplacementRAOs = vdtDisplacementRAOs in requestedData
        self.ImportLoadRAOs = vdtLoadRAOs in requestedData
        self.ImportNewmanQTFs = vdtNewmanQTFs in requestedData
        self.ImportFullQTFs = vdtFullQTFs in requestedData
        self.ImportStiffnessAddedMassDamping = vdtStiffnessAddedMassDamping in requestedData
        self.ImportMassInertia = vdtMassInertia in requestedData
        self.ImportSeaStateRAOs = vdtSeaStateRAOs in requestedData
        self.ImportOtherDamping = vdtOtherDamping in requestedData
        self.WamitLoadRAOMethod = wamitLoadRAOMethods.wlmDefault.Value
        self.WamitDiagonalQTFMethod = wamitDiagonalQTFMethods.wqmDefault.Value
        self.WamitFullQTFMethod = wamitFullQTFMethods.wfmDefault.Value
        if wamitCalculationMethods:
            for method in wamitCalculationMethods:
                if method in wamitLoadRAOMethods:
                    self.WamitLoadRAOMethod = getattr(wamitLoadRAOMethods, method).Value
                if method in wamitDiagonalQTFMethods:
                    self.WamitDiagonalQTFMethod = getattr(wamitDiagonalQTFMethods, method).Value
                if method in wamitFullQTFMethods:
                    self.WamitFullQTFMethod = getattr(wamitFullQTFMethods, method).Value


class VesselTypeDataGenericImportBodyMap(PackedStructure):

    _fields_ = [
        ('DraughtName', _pchar),
        ('RefOrigin', ctypes.c_double * 3),
        ('RefPhaseOrigin', ctypes.c_double * 3)
    ]

    def __init__(
        self,
        DraughtName,
        refOrigin=(
            0.0,
            0.0,
            0.0),
        refPhaseOrigin=(
            OrcinaDefaultReal(),
            OrcinaDefaultReal(),
            OrcinaDefaultReal())):
        PackedStructure.__init__(self)
        self.DraughtName = _PrepareString(DraughtName)
        self.RefOrigin[:] = Vector(refOrigin)
        self.RefPhaseOrigin[:] = Vector(refPhaseOrigin)


class VesselTypeDataGenericImportBodyMapSpecification(PackedStructure):

    _fields_ = [
        ('DestinationVesselType', _pchar),
        ('BodyMapList', ctypes.POINTER(VesselTypeDataGenericImportBodyMap))
    ]

    def __init__(self, DestinationVesselType, bodyCount):
        PackedStructure.__init__(self)
        self.DestinationVesselType = DestinationVesselType
        self.BodyMapList = (VesselTypeDataGenericImportBodyMap * bodyCount)()


class VesselTypeDataDiffractionImportBodyMap(PackedStructure):

    _fields_ = [
        ('Source', _pchar),
        ('DestinationVesselType', _pchar),
        ('DestinationDraught', _pchar)
    ]

    def __init__(self, DestinationVesselType, DestinationDraught):
        PackedStructure.__init__(self)
        self.DestinationVesselType = _PrepareString(DestinationVesselType)
        self.DestinationDraught = _PrepareString(DestinationDraught)


class Interval(PackedStructure):

    _fields_ = [
        ('Lower', ctypes.c_double),
        ('Upper', ctypes.c_double)
    ]


class ExtremeStatisticsSpecification(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('Distribution', ctypes.c_int),
        ('ExtremesToAnalyse', ctypes.c_int),
        ('Threshold', ctypes.c_double),
        ('DeclusterPeriod', ctypes.c_double)
    ]


def RayleighStatisticsSpecification(ExtremesToAnalyse=exUpperTail):
    result = ExtremeStatisticsSpecification()
    result.Distribution = evdRayleigh
    result.ExtremesToAnalyse = ExtremesToAnalyse
    return result


def LikelihoodStatisticsSpecification(Distribution, Threshold, DeclusterPeriod, ExtremesToAnalyse=exUpperTail):
    result = ExtremeStatisticsSpecification()
    result.Distribution = Distribution
    result.Threshold = Threshold
    result.DeclusterPeriod = DeclusterPeriod
    result.ExtremesToAnalyse = ExtremesToAnalyse
    return result


class ExtremeStatisticsQuery(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('StormDurationHours', ctypes.c_double),
        ('RiskFactor', ctypes.c_double),
        ('ConfidenceLevel', ctypes.c_double)
    ]


def RayleighStatisticsQuery(StormDurationHours, RiskFactor):
    result = ExtremeStatisticsQuery()
    result.StormDurationHours = StormDurationHours
    result.RiskFactor = RiskFactor
    return result


def LikelihoodStatisticsQuery(StormDurationHours, ConfidenceLevel):
    result = ExtremeStatisticsQuery()
    result.StormDurationHours = StormDurationHours
    result.ConfidenceLevel = ConfidenceLevel
    return result


class ExtremeStatisticsOutput(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('MostProbableExtremeValue', ctypes.c_double),
        ('ExtremeValueWithRiskFactor', ctypes.c_double),
        ('ReturnLevel', ctypes.c_double),
        ('ConfidenceInterval', Interval),
        ('Sigma', ctypes.c_double),
        ('SigmaStdError', ctypes.c_double),
        ('Xi', ctypes.c_double),
        ('XiStdError', ctypes.c_double)
    ]


class ModalAnalysisSpecification(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('CalculateShapes', BOOL),
        ('FirstMode', ctypes.c_int),
        ('LastMode', ctypes.c_int)
    ]

    def __init__(self, calculateShapes=True, firstMode=-1, lastMode=-1):
        PackedStructure.__init__(self)
        self.CalculateShapes = calculateShapes
        self.FirstMode = firstMode
        self.LastMode = lastMode


class ModeDetails(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('ModeNumber', ctypes.c_int),
        ('Period', ctypes.c_double),
        ('_ShapeWrtGlobal', ctypes.c_void_p),
        ('_ShapeWrtLocal', ctypes.c_void_p)
    ]
    if _checkVersion('9.7a36'):
        _fields_.extend([
            ('ModeType', ctypes.c_int),
            ('PercentageInInlineDirection', ctypes.c_double),
            ('PercentageInAxialDirection', ctypes.c_double),
            ('PercentageInTransverseDirection', ctypes.c_double),
            ('PercentageInRotationalDirection', ctypes.c_double)
        ])

    def __init__(self, dofCount=None):
        PackedStructure.__init__(self)
        if dofCount is None:
            self._ShapeWrtGlobal = None
            self._ShapeWrtLocal = None
        else:
            self.ShapeWrtGlobal = (ctypes.c_double * dofCount)()
            self.ShapeWrtLocal = (ctypes.c_double * dofCount)()
            self._ShapeWrtGlobal = ctypes.cast(ctypes.byref(self.ShapeWrtGlobal), ctypes.c_void_p)
            self._ShapeWrtLocal = ctypes.cast(ctypes.byref(self.ShapeWrtLocal), ctypes.c_void_p)


class SolveEquationParameters(PackedStructure):

    _fields_ = [
        ('Size', ctypes.c_int),
        ('MaxNumberOfIterations', ctypes.c_int),
        ('Tolerance', ctypes.c_double),
        ('MaxStep', ctypes.c_double),
        ('Delta', ctypes.c_double)
    ]

    def __init__(self):
        PackedStructure.__init__(self)
        status = ctypes.c_int()
        _lib.C_GetDefaultSolveEquationParameters(
            ctypes.byref(self),
            ctypes.byref(status)
        )
        _CheckStatus(status)


class GraphCurve(object):

    def __init__(self, X, Y):
        self.X = array(X)
        self.Y = array(Y)

    def __len__(self):
        return 2

    def __getitem__(self, index):
        if index == 0:
            return self.X
        if index == 1:
            return self.Y
        raise IndexError(index)

if NumpyAvailable:

    DoubleArrayType = numpy.ctypeslib.ndpointer(dtype=numpy.float64, ndim=1, flags='C_CONTIGUOUS')

    # http://stackoverflow.com/q/32120178/505088
    _ComplexArrayTypeBase = numpy.ctypeslib.ndpointer(dtype=numpy.complex128, ndim=1, flags='C_CONTIGUOUS')

    def _from_param(cls, obj):
        if obj is None:
            return obj
        return _ComplexArrayTypeBase.from_param(obj)

    ComplexArrayType = type(
        'ComplexArrayType',
        (_ComplexArrayTypeBase,),
        {'from_param': classmethod(_from_param)}
    )

    if hasattr(_lib, 'C_GetFrequencyDomainResultsProcessW'):
        _lib.C_GetFrequencyDomainResultsProcessW.argtypes = (
            Handle,
            ctypes.POINTER(ObjectExtra),
            ctypes.c_int,
            ctypes.POINTER(ctypes.c_int),
            ComplexArrayType,
            ctypes.POINTER(ctypes.c_int)
        )

    if hasattr(_lib, 'C_GetFrequencyDomainResultsFromProcess'):
        _lib.C_GetFrequencyDomainResultsFromProcess.argtypes = (
            Handle,
            ctypes.c_int,
            ComplexArrayType,
            ctypes.POINTER(FrequencyDomainResults),
            ctypes.POINTER(ctypes.c_int)
        )

    if hasattr(_lib, 'C_GetFrequencyDomainSpectralDensityGraphFromProcess'):
        _lib.C_GetFrequencyDomainSpectralDensityGraphFromProcess.argtypes = (
            Handle,
            ctypes.c_int,
            ComplexArrayType,
            ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(_GraphCurve),
            ctypes.POINTER(ctypes.c_int)
        )

    if hasattr(_lib, 'C_GetFrequencyDomainSpectralResponseGraphFromProcess'):
        _lib.C_GetFrequencyDomainSpectralResponseGraphFromProcess.argtypes = (
            Handle,
            ctypes.c_int,
            ComplexArrayType,
            ctypes.POINTER(ctypes.c_int),
            ctypes.POINTER(_GraphCurve),
            ctypes.POINTER(ctypes.c_int)
        )

    if hasattr(_lib, 'C_CalculateMooringStiffness'):
        _lib.C_CalculateMooringStiffness.argtypes = (
            ctypes.c_int,
            ctypes.POINTER(Handle),
            DoubleArrayType,
            ctypes.POINTER(ctypes.c_int)
        )


def DisableModule(Module):
    status = ctypes.c_int()
    _lib.C_DisableModule(
        Module,
        ctypes.byref(status)
    )
    _CheckStatus(status)


def _doFileOperation(handle, operation, filename):
    filename = _PrepareString(filename)
    status = ctypes.c_int()
    _fn('C_' + operation)(handle, _pchar(filename), ctypes.byref(status))
    _CheckStatus(status)


class Model(object):

    _hookOrcaFlexObjectClass = []

    @classmethod
    def addOrcaFlexObjectClassHook(cls, hook):
        cls._hookOrcaFlexObjectClass.append(hook)

    @classmethod
    def removeOrcaFlexObjectClassHook(cls, hook):
        cls._hookOrcaFlexObjectClass.remove(hook)

    def __init__(self, filename=None, threadCount=None, handle=None):
        status = ctypes.c_int()
        self.ownsModelHandle = not handle
        if self.ownsModelHandle:
            handle = Handle()
            if _CreateModel2Available and not threadCount is None:
                params = CreateModelParams()
                params.ThreadCount = threadCount
                _lib.C_CreateModel2(ctypes.byref(handle), ctypes.byref(params), ctypes.byref(status))
                _CheckStatus(status)
                self.handle = handle
            else:
                _lib.C_CreateModel(ctypes.byref(handle), 0, ctypes.byref(status))
                _CheckStatus(status)
                self.handle = handle
                if threadCount is not None:
                    self.threadCount = threadCount
        else:
            self.handle = handle
        self.general = self['General']
        self.environment = self['Environment']
        self.progressHandler = None
        self.batchProgressHandler = None
        self.staticsProgressHandler = None
        self.dynamicsProgressHandler = None
        self.postCalculationActionProgressHandler = None
        # avoid calling C_SetCorrectExternalFileReferencesHandler which fails on
        # older versions of the DLL
        object.__setattr__(self, 'correctExternalFileReferencesHandler', None)
        if filename:
            filename = _PrepareString(filename)
            status = ctypes.c_int()
            _fn('C_LoadSimulation')(handle, _pchar(filename), ctypes.byref(status))
            if status.value != stOK:
                _fn('C_LoadData')(handle, _pchar(filename), ctypes.byref(status))
                _CheckStatus(status)

    def __del__(self):
        if self.ownsModelHandle:
            try:
                status = ctypes.c_int()
                _lib.C_DestroyModel(self.handle, ctypes.byref(status))
                # no point checking status here since we can't really do anything about a failure
            except:
                # swallow this since we get exceptions when Python terminates unexpectedly
                # (e.g. CTRL+Z)
                pass

    def __setattr__(self, name, value):
        object.__setattr__(self, name, value)
        status = ctypes.c_int()
        if name == 'progressHandler':
            if self.progressHandler:
                def callback(handle, progress, cancel):
                    cancel[0] = _ProgressCallbackCancel(self.progressHandler(self, progress))
                self.progressHandlerCallback = ctypes.WINFUNCTYPE(
                    None, Handle, ctypes.c_int, ctypes.POINTER(BOOL))(callback)
            else:
                self.progressHandlerCallback = None
            _lib.C_SetProgressHandler(
                self.handle,
                self.progressHandlerCallback,
                ctypes.byref(status)
            )
            _CheckStatus(status)
        elif name == 'correctExternalFileReferencesHandler':
            if self.correctExternalFileReferencesHandler:
                def callback(handle):
                    self.correctExternalFileReferencesHandler(self)
                self.correctExternalFileReferencesHandlerCallback = ctypes.WINFUNCTYPE(
                    None, Handle)(callback)
            else:
                self.correctExternalFileReferencesHandlerCallback = None
            _lib.C_SetCorrectExternalFileReferencesHandler(
                self.handle,
                self.correctExternalFileReferencesHandlerCallback,
                ctypes.byref(status)
            )
            _CheckStatus(status)

    def __getitem__(self, name):
        objectInfo = ObjectInfo()
        name = _PrepareString(name)
        status = ctypes.c_int()
        _fn('C_ObjectCalled')(
            self.handle,
            _pchar(name),
            ctypes.byref(objectInfo),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return self.createOrcaFlexObject(Handle(objectInfo.ObjectHandle), objectInfo.ObjectType)

    def _stringProgressHandler(self, handler):
        if handler:
            def callback(handle, progress, cancel):
                cancel[0] = _ProgressCallbackCancel(handler(self, progress))
            return ctypes.WINFUNCTYPE(None, Handle, _pchar, ctypes.POINTER(BOOL))(callback)
        else:
            return None

    def _batchProgressHandler(self):
        return self._stringProgressHandler(self.batchProgressHandler)

    def _staticsProgressHandler(self):
        return self._stringProgressHandler(self.staticsProgressHandler)

    def _dynamicsProgressHandler(self):
        if self.dynamicsProgressHandler:
            def callback(handle, time, start, stop, cancel):
                cancel[0] = _ProgressCallbackCancel(
                    self.dynamicsProgressHandler(self, time, start, stop))
            return ctypes.WINFUNCTYPE(
                None,
                Handle,
                ctypes.c_double,
                ctypes.c_double,
                ctypes.c_double,
                ctypes.POINTER(BOOL))(callback)
        else:
            return None

    def _postCalculationActionProgressHandler(self):
        return self._stringProgressHandler(self.postCalculationActionProgressHandler)

    def orcaFlexObjectClass(self, type):
        if type == otLine:
            objectClass = OrcaFlexLineObject
        elif type == otLineType or type == otBendingStiffness:
            objectClass = OrcaFlexWizardObject
        elif type == otVessel:
            objectClass = OrcaFlexVesselObject
        else:
            objectClass = OrcaFlexObject
        for hook in Model._hookOrcaFlexObjectClass:
            objectClass = hook(type, objectClass)
        return objectClass

    def createOrcaFlexObject(self, handle, type=None):
        if type is None:
            name = _PrepareString('Name')
            return self[_GetDataString(handle, _pchar(name).value, -1)]
        else:
            return self.orcaFlexObjectClass(type)(self.handle, handle, type)

    def ModuleEnabled(self, module):
        status = ctypes.c_int()
        result = _lib.C_ModuleEnabled(
            self.handle,
            module,
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result != 0

    @property
    def threadCount(self):
        status = ctypes.c_int()
        result = _lib.C_GetModelThreadCount(self.handle, ctypes.byref(status))
        _CheckStatus(status)
        return result

    @threadCount.setter
    def threadCount(self, value):
        status = ctypes.c_int()
        _lib.C_SetModelThreadCount(self.handle, ctypes.c_int(value), ctypes.byref(status))
        _CheckStatus(status)

    @property
    def recommendedTimeSteps(self):
        result = TimeSteps()
        status = ctypes.c_int()
        _lib.C_GetRecommendedTimeSteps(
            self.handle,
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result

    @property
    def simulationStartTime(self):
        return self.simulationTimeStatus.StartTime

    @property
    def simulationStopTime(self):
        return self.simulationTimeStatus.StopTime

    @property
    def simulationTimeToGo(self):
        status = ctypes.c_int()
        result = _lib.C_GetSimulationTimeToGo(self.handle, ctypes.byref(status))
        _CheckStatus(status)
        return result

    @property
    def state(self):
        state = ctypes.c_int()
        status = ctypes.c_int()
        _lib.C_GetModelState(self.handle, ctypes.byref(state), ctypes.byref(status))
        _CheckStatus(status)
        return ModelState[state.value]

    @property
    def simulationComplete(self):
        simulationComplete = ctypes.c_int()
        status = ctypes.c_int()
        _lib.C_GetSimulationComplete(self.handle, ctypes.byref(simulationComplete), ctypes.byref(status))
        _CheckStatus(status)
        return simulationComplete.value != 0

    @property
    def simulationTimeStatus(self):
        result = SimulationTimeStatus()
        status = ctypes.c_int()
        _lib.C_GetSimulationTimeStatus(
            self.handle,
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result

    def getBooleanModelProperty(self, propertyId):
        result = BOOL()
        status = ctypes.c_int()
        _lib.C_GetModelProperty(
            self.handle,
            propertyId,
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return bool(result)

    @property
    def isTimeDomainDynamics(self):
        return not _checkVersion('9.9a34') or self.getBooleanModelProperty(propIsTimeDomainDynamics)

    @property
    def isFrequencyDomainDynamics(self):
        return _checkVersion('9.9a34') and self.getBooleanModelProperty(propIsFrequencyDomainDynamics)

    @property
    def isDeterministicFrequencyDomainDynamics(self):
        return _checkVersion('9.9a34') and self.getBooleanModelProperty(propIsDeterministicFrequencyDomainDynamics)

    @property
    def objects(self):
        result = []

        def callback(handle, objectInfo):
            result.append(self.createOrcaFlexObject(
                Handle(objectInfo[0].ObjectHandle), objectInfo[0].ObjectType))
        status = ctypes.c_int()
        _fn('C_EnumerateObjects')(
            self.handle,
            ctypes.WINFUNCTYPE(None, Handle, ctypes.POINTER(ObjectInfo))(callback),
            ctypes.byref(ctypes.c_int()),  # NumOfObjects
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return tuple(result)

    def CreateObject(self, type, name=None):
        handle = Handle()
        type = int(type)
        status = ctypes.c_int()
        _lib.C_CreateObject(
            self.handle,
            type,
            ctypes.byref(handle),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        obj = self.createOrcaFlexObject(handle, type)
        if name:
            obj.name = name
        return obj

    def DestroyObject(self, obj):
        if not hasattr(obj, 'handle'):
            obj = self[obj]  # assume that obj is a name
        status = ctypes.c_int()
        _lib.C_DestroyObject(obj.handle, ctypes.byref(status))
        _CheckStatus(status)

    def Reset(self):
        status = ctypes.c_int()
        _lib.C_ResetModel(self.handle, ctypes.byref(status))
        _CheckStatus(status)

    def Clear(self):
        status = ctypes.c_int()
        _lib.C_ClearModel(self.handle, ctypes.byref(status))
        _CheckStatus(status)

    def LoadData(self, filename):
        _doFileOperation(self.handle, 'LoadData', filename)

    def SaveData(self, filename):
        _doFileOperation(self.handle, 'SaveData', filename)

    def LoadSimulation(self, filename):
        _doFileOperation(self.handle, 'LoadSimulation', filename)

    def SaveSimulation(self, filename):
        _doFileOperation(self.handle, 'SaveSimulation', filename)

    @property
    def warnings(self):

        def warningGenerator():
            status = ctypes.c_int()
            count = _lib.C_GetNumOfWarnings(self.handle, ctypes.byref(status))
            _CheckStatus(status)
            func = _fn('C_GetWarningText')
            for index in range(count):
                length = func(
                    self.handle,
                    ctypes.c_int(index),
                    None,
                    None,
                    ctypes.byref(status)
                )
                _CheckStatus(status)
                warningText = (_char * length)()
                func(
                    self.handle,
                    ctypes.c_int(index),
                    None,
                    ctypes.byref(warningText),
                    ctypes.byref(status)
                )
                _CheckStatus(status)
                yield warningText.value

        return tuple(warningGenerator())

    @property
    def waveComponents(self):

        def componentGenerator():
            count = ctypes.c_int()
            status = ctypes.c_int()
            if _GetWaveComponents2Available:
                GetWaveComponents = _lib.C_GetWaveComponents2
            else:
                GetWaveComponents = _lib.C_GetWaveComponents
            GetWaveComponents(
                self.handle,
                ctypes.byref(count),
                None,
                ctypes.byref(status)
            )
            _CheckStatus(status)

            components = (WaveComponent * count.value)()
            GetWaveComponents(
                self.handle,
                ctypes.byref(count),
                ctypes.byref(components),
                ctypes.byref(status)
            )
            _CheckStatus(status)

            for index in range(count.value):
                yield components[index].asObject()

        return tuple(componentGenerator())

    @property
    def windComponents(self):

        def componentGenerator():
            count = ctypes.c_int()
            status = ctypes.c_int()
            _lib.C_GetWindComponents(
                self.handle,
                ctypes.byref(count),
                None,
                ctypes.byref(status)
            )
            _CheckStatus(status)

            components = (WindComponent * count.value)()
            _lib.C_GetWindComponents(
                self.handle,
                ctypes.byref(count),
                ctypes.byref(components),
                ctypes.byref(status)
            )
            _CheckStatus(status)

            for index in range(count.value):
                yield components[index].asObject()

        return tuple(componentGenerator())

    def DisableInMemoryLogging(self):
        status = ctypes.c_int()
        _lib.C_DisableInMemoryLogging(
            self.handle,
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def UseVirtualLogging(self):
        status = ctypes.c_int()
        _lib.C_UseVirtualLogging(
            self.handle,
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def CalculateStatics(self):
        status = ctypes.c_int()
        _fn('C_CalculateStatics')(
            self.handle,
            self._staticsProgressHandler(),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def RunSimulation(
            self,
            enableAutoSave=False,
            autoSaveIntervalMinutes=DefaultAutoSaveIntervalMinutes,
            autoSaveFileName=None):
        status = ctypes.c_int()
        if self.state <= ModelState.Reset:
            self.CalculateStatics()
        _fn('C_RunSimulation2')(
            self.handle,
            self._dynamicsProgressHandler(),
            ctypes.byref(RunSimulationParameters(
                enableAutoSave, autoSaveIntervalMinutes, autoSaveFileName)),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def ProcessBatchScript(
            self,
            filename,
            enableAutoSave=False,
            autoSaveIntervalMinutes=DefaultAutoSaveIntervalMinutes,
            autoSaveFileName=None):
        filename = _PrepareString(filename)
        status = ctypes.c_int()
        _fn('C_ProcessBatchScript')(
            self.handle,
            _pchar(filename),
            self._batchProgressHandler(),
            self._staticsProgressHandler(),
            self._dynamicsProgressHandler(),
            ctypes.byref(RunSimulationParameters(
                enableAutoSave, autoSaveIntervalMinutes, autoSaveFileName)),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def ExtendSimulation(self, time):
        time = ctypes.c_double(time)
        status = ctypes.c_int()
        _lib.C_ExtendSimulation(
            self.handle,
            time,
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def PauseSimulation(self):
        if self.state == ModelState.RunningSimulation:
            status = ctypes.c_int()
            _lib.C_PauseSimulation(
                self.handle,
                ctypes.byref(status)
            )
            _CheckStatus(status)

    def InvokeLineSetupWizard(self):
        status = ctypes.c_int()
        _fn('C_InvokeLineSetupWizard')(
            self.handle,
            self._staticsProgressHandler(),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def UseCalculatedPositions(self, SetLinesToUserSpecifiedStartingShape=False):
        Parameters = UseCalculatedPositionsForStaticsParameters()
        Parameters.SetLinesToUserSpecifiedStartingShape = SetLinesToUserSpecifiedStartingShape
        status = ctypes.c_int()
        _lib.C_UseCalculatedPositionsForStatics(
            self.handle,
            ctypes.byref(Parameters),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    @property
    def defaultViewParameters(self):
        result = ViewParameters()
        status = ctypes.c_int()
        _lib.C_GetDefaultViewParameters(
            self.handle,
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result

    def checkViewParameters(self, viewParameters):
        if viewParameters is not None:
            if not isinstance(viewParameters, ViewParameters):
                raise TypeError(
                    'viewParameters parameter must None, or be an instance of class ViewParameters')
            viewParameters = ctypes.byref(viewParameters)
        return viewParameters

    def SaveModelView(self, filename, viewParameters=None):
        filename = _PrepareString(filename)
        status = ctypes.c_int()
        _fn('C_SaveModel3DViewBitmapToFile')(
            self.handle,
            self.checkViewParameters(viewParameters),
            _pchar(filename),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    @property
    def simulationDrawTime(self):
        status = ctypes.c_int()
        result = _lib.C_GetSimulationDrawTime(
            self.handle,
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result

    @simulationDrawTime.setter
    def simulationDrawTime(self, value):
        status = ctypes.c_int()
        _lib.C_SetSimulationDrawTime(
            self.handle,
            ctypes.c_double(value),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def SaveWaveSearchSpreadsheet(self, filename):
        _SaveSpreadsheet(self.handle, sptWaveSearch, filename, None)

    def SaveLineTypesPropertiesSpreadsheet(self, filename):
        _SaveSpreadsheet(self.handle, sptLineTypesPropertiesReport, filename, None)

    def SaveCodeChecksProperties(self, filename):
        _SaveSpreadsheet(self.handle, sptCodeChecksProperties, filename, None)

    def createPeriodParameter(self, period):
        if period is None:
            if self.state == ModelState.InStaticState:
                period = Period(pnStaticState)
            else:
                period = Period(pnWholeSimulation)
        elif not isinstance(period, Period):
            if isinstance(period, int):
                period = Period(period)
            else:
                period = SpecifiedPeriod(period[0], period[1])
        return period

    def checkPeriod(self, period):
        return ctypes.byref(self.createPeriodParameter(period))

    def SampleTimes(self, period=None):
        status = ctypes.c_int()
        sampleCount = _lib.C_GetNumOfSamples(
            self.handle,
            self.checkPeriod(period),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        samples = (ctypes.c_double * sampleCount)()
        _lib.C_GetSampleTimes(
            self.handle,
            self.checkPeriod(period),
            ctypes.byref(samples),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return array(samples)

    def ExecutePostCalculationActions(self, filename, actionType, treatExecutionErrorsAsWarnings=False):
        filename = _PrepareString(filename)
        status = ctypes.c_int()
        _fn('C_ExecutePostCalculationActions')(
            self.handle,
            _pchar(filename),
            self._postCalculationActionProgressHandler(),
            ctypes.c_int(actionType),
            BOOL(treatExecutionErrorsAsWarnings),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def ReportActionProgress(self, progress):
        if _ReportActionProgressAvailable:
            progress = _PrepareString(progress)
            status = ctypes.c_int()
            _lib.C_ReportActionProgressW(self.handle, _pchar(progress).value, ctypes.byref(status))
            _CheckStatus(status)

    # Unsupported, undocumented, internal function, do not call it
    def ImportVesselTypeData(
        self,
        filename,
        fileType,
        mappings,
        wamitCalculationMethods=None,
        destVesselType='',
        multibodyGroup='',
        requestedData=()
    ):
        try:
            bodyCount = len(mappings)
        except TypeError:
            bodyCount = 1
            mappings = [mappings]
        multibodyGroup = _PrepareString(multibodyGroup)
        specification = VesselTypeDataImportSpecification(
            fileType, bodyCount, multibodyGroup, requestedData, wamitCalculationMethods)

        def prepareGenericBodyMap(destVesselType, count, mappings):
            bodyMap = VesselTypeDataGenericImportBodyMapSpecification(destVesselType, count)
            for i, map in enumerate(mappings):
                bodyMap.BodyMapList[i] = map
            return bodyMap

        def prepareDiffractionBodyMap(fileType, count, mappings):
            bodyMap = (VesselTypeDataDiffractionImportBodyMap * count)()
            for i, map in enumerate(mappings):
                bodyMap[i] = map
                if fileType == iftAQWA:
                    bodyMap[i].Source = 'Structure %d' % (i + 1)
                elif fileType == iftWAMIT:
                    bodyMap[i].Source = 'Body number %d' % (i + 1)
            return bodyMap

        if fileType == iftGeneric:
            bodyMap = prepareGenericBodyMap(destVesselType, bodyCount, mappings)
        elif fileType in [iftAQWA, iftWAMIT]:
            bodyMap = prepareDiffractionBodyMap(fileType, bodyCount, mappings)
        else:
            raise TypeError('invalid dataType: ', fileType)
        messages = ctypes.c_void_p()
        filename = _PrepareString(filename)
        status = ctypes.c_int()

        _lib.C_ImportVesselTypeDataW(
            self.handle,
            _pchar(filename),
            ctypes.byref(specification),
            ctypes.byref(bodyMap),
            ctypes.byref(messages),
            ctypes.byref(status)
        )

        def decode(messages):
            # have to perform manual marshalling because memory is allocated in DLL
            # and not by ctypes
            bufferLength = ctypes.cdll.msvcrt.wcslen(messages) + 1
            buffer = (_char * bufferLength)()
            ctypes.cdll.msvcrt.wcsncpy(ctypes.byref(buffer), messages, bufferLength)
            return _DecodeString(buffer.value)
        if status.value == stVesselTypeDataImportFailed:
            result = False, decode(messages)
        else:
            _CheckStatus(status)
            result = True, decode(messages)
        ctypes.windll.OleAut32.SysFreeString(messages.value)
        return result

    def FrequencyDomainResultsFromProcess(self, process):
        _requiresNumpy()
        if not process.flags.c_contiguous:
            process = numpy.ascontiguousarray(process, dtype=numpy.complex128)
        result = FrequencyDomainResults()
        status = ctypes.c_int()
        _lib.C_GetFrequencyDomainResultsFromProcess(self.handle, len(process), process, result, status)
        _CheckStatus(status)
        return result.asObject()

    def FrequencyDomainSpectralDensityFromProcess(self, process):
        _requiresNumpy()
        if not process.flags.c_contiguous:
            process = numpy.ascontiguousarray(process, dtype=numpy.complex128)
        pointCount = ctypes.c_int()
        status = ctypes.c_int()
        func = _lib.C_GetFrequencyDomainSpectralDensityGraphFromProcess
        func(self.handle, len(process), None, pointCount, None, status)
        _CheckStatus(status)
        curve = _GraphCurve(pointCount.value)
        func(self.handle, len(process), process, pointCount, curve, status)
        _CheckStatus(status)
        return GraphCurve(curve.X, curve.Y)

    def FrequencyDomainSpectralResponseRAOFromProcess(self, process):
        _requiresNumpy()
        if not process.flags.c_contiguous:
            process = numpy.ascontiguousarray(process, dtype=numpy.complex128)
        pointCount = ctypes.c_int()
        status = ctypes.c_int()
        func = _lib.C_GetFrequencyDomainSpectralResponseGraphFromProcess
        func(self.handle, len(process), None, pointCount, None, status)
        _CheckStatus(status)
        curve = _GraphCurve(pointCount.value)
        func(self.handle, len(process), process, pointCount, curve, status)
        _CheckStatus(status)
        return GraphCurve(curve.X, curve.Y)


class DataObject(object):

    def __init__(self, handle):
        self.handle = handle

    def __eq__(self, other):
        return (isinstance(self, type(other))) and (self.handle.value == other.handle.value)

    def __ne__(self, other):
        return not self.__eq__(other)

    def isBuiltInAttribute(self, name):
        return name in ('handle', 'status')

    def __getattr__(self, name):
        if self.isBuiltInAttribute(name):
            return object.__getattr__(self, name)
        elif self.DataNameValid(name):
            if self.DataNameRequiresIndex(name):
                return IndexedDataItem(name, self)
            else:
                return self.GetData(name, -1)
        else:
            raise AttributeError(name)

    def __setattr__(self, name, value):
        if self.isBuiltInAttribute(name):
            object.__setattr__(self, name, value)
        elif self.DataNameValid(name):
            if self.DataNameRequiresIndex(name):
                IndexedDataItem(name, self).Assign(value)
            else:
                self.SetData(name, -1, value)
        else:
            raise AttributeError(name)

    def DataType(self, dataName):
        dataName = _PrepareString(dataName)
        status = ctypes.c_int()
        result = ctypes.c_int()
        _fn('C_GetDataType')(
            self.handle,
            _pchar(dataName),
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.value

    def VariableDataType(self, dataName, index):
        index = int(index) + 1  # Python is 0-based, OrcFxAPI is 1-based
        dataName = _PrepareString(dataName)
        status = ctypes.c_int()
        result = ctypes.c_int()
        _fn('C_GetVariableDataType')(
            self.handle,
            _pchar(dataName),
            index,
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.value

    def DataNameValid(self, dataName):
        dataName = _PrepareString(dataName)
        status = ctypes.c_int()
        _fn('C_GetDataType')(
            self.handle,
            _pchar(dataName),
            ctypes.byref(ctypes.c_int()),  # DataType
            ctypes.byref(status)
        )
        return status.value == stOK

    def DataNameRequiresIndex(self, dataName):
        if _DataCountAvailable:
            return self.GetDataRowCount(dataName) != -1
        else:
            result = False
            try:
                self.GetData(name, -1)
            except DLLError:
                result = True
            return result

    def GetDataRowCount(self, indexedDataName):
        if _DataCountAvailable:
            rowCount = ctypes.c_int()
            indexedDataName = _PrepareString(indexedDataName)
            status = ctypes.c_int()
            _fn('C_GetDataRowCount')(
                self.handle,
                _pchar(indexedDataName),
                ctypes.byref(rowCount),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            return rowCount.value
        else:
            result = 0
            try:
                while True:
                    self.GetData(indexedDataName, result)
                    result += 1
            except:
                return result

    def SetDataRowCount(self, indexedDataName, rowCount):
        indexedDataName = _PrepareString(indexedDataName)
        status = ctypes.c_int()
        _fn('C_SetDataRowCount')(
            self.handle,
            _pchar(indexedDataName),
            ctypes.c_int(rowCount),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def InsertDataRow(self, indexedDataName, index):
        index = int(index) + 1  # Python is 0-based, OrcFxAPI is 1-based
        indexedDataName = _PrepareString(indexedDataName)
        status = ctypes.c_int()
        _fn('C_InsertDataRow')(
            self.handle,
            _pchar(indexedDataName),
            ctypes.c_int(index),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def DeleteDataRow(self, indexedDataName, index):
        index = int(index) + 1  # Python is 0-based, OrcFxAPI is 1-based
        indexedDataName = _PrepareString(indexedDataName)
        status = ctypes.c_int()
        _fn('C_DeleteDataRow')(
            self.handle,
            _pchar(indexedDataName),
            ctypes.c_int(index),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def dataChange(self):

        class dataChangeContext(object):

            def __init__(self, obj):
                self.obj = obj

            def __enter__(self):
                if _BeginEndDataChangeAvailable:
                    status = ctypes.c_int()
                    _lib.C_BeginDataChange(
                        self.obj.handle,
                        ctypes.byref(status)
                    )
                    _CheckStatus(status)

            def __exit__(self, type, value, traceback):
                if _BeginEndDataChangeAvailable:
                    status = ctypes.c_int()
                    _lib.C_EndDataChange(
                        self.obj.handle,
                        ctypes.byref(status)
                    )
                    _CheckStatus(status)

        return dataChangeContext(self)

    def GetData(self, dataNames, index):

        def GetSingleDataItem(dataName, index):
            index = int(index) + 1  # Python is 0-based, OrcFxAPI is 1-based
            dataType = ctypes.c_int()
            dataName = _PrepareString(dataName)
            status = ctypes.c_int()
            _fn('C_GetDataType')(
                self.handle,
                _pchar(dataName),
                ctypes.byref(dataType),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            if dataType.value == dtVariable:
                _fn('C_GetVariableDataType')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    ctypes.byref(dataType),
                    ctypes.byref(status)
                )
                if status.value == stValueNotAvailable:
                    return None
                _CheckStatus(status)
            if dataType.value == dtDouble:
                result = ctypes.c_double()
                _fn('C_GetDataDouble')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    ctypes.byref(result),
                    ctypes.byref(status)
                )
                if status.value == stValueNotAvailable:
                    return None
                _CheckStatus(status)
                return result.value
            elif dataType.value == dtInteger:
                result = ctypes.c_int()
                _fn('C_GetDataInteger')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    ctypes.byref(result),
                    ctypes.byref(status)
                )
                if status.value == stValueNotAvailable:
                    return None
                _CheckStatus(status)
                return result.value
            elif dataType.value == dtString:
                return _GetDataString(self.handle, dataName, index)
            else:
                raise DLLError('Unrecognised data type')

        if isinstance(dataNames, str if _isPy3k else basestring):
            return GetSingleDataItem(dataNames, index)
        else:
            return [GetSingleDataItem(dataName, index) for dataName in dataNames]

    def SetData(self, dataNames, index, values):

        def SetSingleDataItem(dataName, index, value):
            dataName = _PrepareString(dataName)
            index = int(index) + 1  # Python is 0-based, OrcFxAPI is 1-based
            dataType = ctypes.c_int()
            status = ctypes.c_int()
            _fn('C_GetDataType')(
                self.handle,
                _pchar(dataName),
                ctypes.byref(dataType),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            if dataType.value == dtVariable:
                try:
                    float(value)
                    dataType.value = dtDouble
                except:
                    dataType.value = dtString
            if dataType.value == dtDouble:
                value = ctypes.c_double(value)
                _fn('C_SetDataDouble')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    value,
                    ctypes.byref(status)
                )
                _CheckStatus(status)
            elif dataType.value == dtInteger:
                value = ctypes.c_int(value)
                _fn('C_SetDataInteger')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    value,
                    ctypes.byref(status)
                )
                _CheckStatus(status)
            elif dataType.value == dtString:
                if isinstance(value, bool):
                    value = str(value)
                value = _PrepareString(value)
                _fn('C_SetDataString')(
                    self.handle,
                    _pchar(dataName),
                    index,
                    _pchar(value),
                    ctypes.byref(status)
                )
                _CheckStatus(status)
            else:
                raise DLLError('Unrecognised data type')

        if isinstance(dataNames, str if _isPy3k else basestring):
            SetSingleDataItem(dataNames, index, values)
        else:
            [SetSingleDataItem(dataName, index, value)
             for dataName, value in zip(dataNames, values)]


class OrcaFlexObject(DataObject):

    def __init__(self, modelHandle, handle, type):
        DataObject.__init__(self, handle)
        self.modelHandle = modelHandle
        self.type = type

    def isBuiltInAttribute(self, name):
        return name in ('modelHandle', 'type') or DataObject.isBuiltInAttribute(self, name)

    def __repr__(self):
        return "<%s: '%s'>" % (self.typeName, self.name)

    def CreateClone(self, name=None, model=None):
        handle = Handle()
        status = ctypes.c_int()
        if model is None:
            _lib.C_CreateClone(
                self.handle,
                ctypes.byref(handle),
                ctypes.byref(status)
            )
        else:
            _lib.C_CreateClone2(
                self.handle,
                model.handle,
                ctypes.byref(handle),
                ctypes.byref(status)
            )
        _CheckStatus(status)
        if model is None:
            clone = self.tempModel().createOrcaFlexObject(handle, self.type)
        else:
            clone = model.createOrcaFlexObject(handle, self.type)
        if name is not None:
            clone.name = name
        return clone

    @property
    def typeName(self):
        func = _fn('C_GetObjectTypeName')
        status = ctypes.c_int()
        length = func(self.modelHandle, self.type, None, ctypes.byref(status))
        _CheckStatus(status)
        name = (_char * length)()
        func(self.modelHandle, self.type, ctypes.byref(name), ctypes.byref(status))
        _CheckStatus(status)
        return _DecodeString(name.value)

    def tempModel(self):
        return Model(handle=self.modelHandle)

    def createPeriodParameter(self, period):
        return self.tempModel().createPeriodParameter(period)

    def checkPeriod(self, period):
        return ctypes.byref(self.createPeriodParameter(period))

    def checkObjectExtra(self, objectExtra):
        if objectExtra is not None:
            if not isinstance(objectExtra, ObjectExtra):
                raise TypeError(
                    'objectExtra parameter must be either None or an instance of class ObjectExtra')
            objectExtra = ctypes.byref(objectExtra)
        return objectExtra

    def SampleTimes(self, period=None):
        return self.tempModel().SampleTimes(period)

    def varID(self, varName):
        varName = _PrepareString(varName)
        varID = ctypes.c_int()
        status = ctypes.c_int()
        _fn('C_GetVarID')(
            self.handle,
            _pchar(varName),
            ctypes.byref(varID),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return varID.value

    def varDetails(self, resultType=rtTimeHistory, objectExtra=None):
        result = []

        def callback(varInfo):
            result.append(
                objectFromDict({
                    'VarName': varInfo[0].VarName,
                    'VarUnits': varInfo[0].VarUnits,
                    'FullName': varInfo[0].FullName
                })
            )
        status = ctypes.c_int()
        _fn('C_EnumerateVars2')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            ctypes.c_int(resultType),
            ctypes.WINFUNCTYPE(None, ctypes.POINTER(VarInfo))(callback),
            ctypes.byref(ctypes.c_int()),  # NumOfVars
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return tuple(result)

    def vars(self, resultType=rtTimeHistory, objectExtra=None):
        result = []

        def callback(varInfo):
            result.append(varInfo[0].VarName)
        status = ctypes.c_int()
        _fn('C_EnumerateVars2')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            ctypes.c_int(resultType),
            ctypes.WINFUNCTYPE(None, ctypes.POINTER(VarInfo))(callback),
            ctypes.byref(ctypes.c_int()),  # NumOfVars
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return tuple(result)

    def TimeHistory(self, varNames, period=None, objectExtra=None):

        def getTimeHistory(varName):
            status = ctypes.c_int()
            sampleCount = _lib.C_GetNumOfSamples(
                self.modelHandle,
                self.checkPeriod(period),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            values = (ctypes.c_double * sampleCount)()
            _fn('C_GetTimeHistory2')(
                self.handle,
                self.checkObjectExtra(objectExtra),
                self.checkPeriod(period),
                ctypes.c_int(self.varID(varName)),
                ctypes.byref(values),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            return array(values)

        if isinstance(varNames, str if _isPy3k else basestring):
            return getTimeHistory(varNames)
        else:
            hist = []
            for varName in varNames:
                hist.append(getTimeHistory(varName))
            return numpy.column_stack(hist)

    def StaticResult(self, varNames, objectExtra=None):
        return self.TimeHistory(varNames, Period(pnStaticState), objectExtra)[0]

    def LinkedStatistics(self, varNames, period=None, objectExtra=None):
        return LinkedStatistics(self, varNames, period, objectExtra)

    def TimeSeriesStatistics(self, varName, period=None, objectExtra=None):
        return TimeSeriesStatistics(
            self.TimeHistory(varName, period, objectExtra),
            self.tempModel().general.ActualLogSampleInterval
        )

    def ExtremeStatistics(self, varName, period=None, objectExtra=None):
        values = self.TimeHistory(varName, period, objectExtra)
        sampleInterval = self.tempModel().general.ActualLogSampleInterval
        return ExtremeStatistics(values, sampleInterval)

    def FrequencyDomainResults(self, varName, objectExtra=None):
        result = FrequencyDomainResults()
        status = ctypes.c_int()
        _fn('C_GetFrequencyDomainResults')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            ctypes.c_int(self.varID(varName)),
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.asObject()

    def FrequencyDomainResultsProcess(self, varName, objectExtra=None):
        _requiresNumpy()
        objectExtra = self.checkObjectExtra(objectExtra)
        varID = self.varID(varName)
        componentCount = ctypes.c_int()
        status = ctypes.c_int()
        func = _lib.C_GetFrequencyDomainResultsProcessW
        func(self.handle, objectExtra, varID, componentCount, None, status)
        _CheckStatus(status)
        process = numpy.zeros(componentCount.value, dtype=numpy.complex128)
        func(self.handle, objectExtra, varID, componentCount, process, status)
        _CheckStatus(status)
        return process

    def FrequencyDomainMPM(self, varName, stormDurationHours, objectExtra=None):
        fdr = self.FrequencyDomainResults(varName, objectExtra)
        return FrequencyDomainMPM(stormDurationHours * 3600.0, fdr.StdDev, fdr.Tz)

    def AnalyseExtrema(self, varName, period=None, objectExtra=None):
        return AnalyseExtrema(
            self.TimeHistory(varName, period, objectExtra)
        )

    def SpectralResponseRAO(self, varName, objectExtra=None):
        varID = ctypes.c_int(self.varID(varName))
        NumOfGraphPoints = ctypes.c_int()
        status = ctypes.c_int()
        _fn('C_GetSpectralResponseGraph')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            varID,
            ctypes.byref(NumOfGraphPoints),
            None,
            ctypes.byref(status)
        )
        _CheckStatus(status)

        curve = _GraphCurve(NumOfGraphPoints.value)
        _fn('C_GetSpectralResponseGraph')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            varID,
            ctypes.byref(NumOfGraphPoints),
            ctypes.byref(curve),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return GraphCurve(curve.X, curve.Y)

    def SpectralDensity(self, varName, period=None, objectExtra=None, fundamentalFrequency=None):
        model = self.tempModel()
        general = model.general
        if model.isFrequencyDomainDynamics:
            varID = ctypes.c_int(self.varID(varName))
            NumOfGraphPoints = ctypes.c_int()
            status = ctypes.c_int()
            _fn('C_GetFrequencyDomainSpectralDensityGraph')(
                self.handle,
                self.checkObjectExtra(objectExtra),
                varID,
                ctypes.byref(NumOfGraphPoints),
                None,
                ctypes.byref(status)
            )
            _CheckStatus(status)

            curve = _GraphCurve(NumOfGraphPoints.value)
            _fn('C_GetFrequencyDomainSpectralDensityGraph')(
                self.handle,
                self.checkObjectExtra(objectExtra),
                varID,
                ctypes.byref(NumOfGraphPoints),
                ctypes.byref(curve),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            return GraphCurve(curve.X, curve.Y)
        else:
            if fundamentalFrequency is None:
                if general.DataNameValid('SpectralDensityFundamentalFrequency'):
                    fundamentalFrequency = general.SpectralDensityFundamentalFrequency
            return SpectralDensity(
                self.SampleTimes(period),
                self.TimeHistory(varName, period, objectExtra),
                fundamentalFrequency
            )

    def EmpiricalDistribution(self, varName, period=None, objectExtra=None):
        return EmpiricalDistribution(
            self.TimeHistory(varName, period, objectExtra)
        )

    def RainflowHalfCycles(self, varName, period=None, objectExtra=None):
        return RainflowHalfCycles(
            self.TimeHistory(varName, period, objectExtra)
        )

    def SaveSummaryResults(self, filename):
        _SaveSpreadsheet(self.handle, sptSummaryResults, filename, None)

    def SaveFullResults(self, filename):
        _SaveSpreadsheet(self.handle, sptFullResults, filename, None)

    def SaveDetailedPropertiesReport(self, filename):
        _SaveSpreadsheet(self.handle, sptDetailedProperties, filename, None)


class OrcaFlexLineObject(OrcaFlexObject):

    def lineTypeAt(self, arclength):
        sectionLength = array(self.CumulativeLength)
        index = 0
        while index < len(sectionLength) - 1:
            if arclength <= sectionLength[index]:
                break
            else:
                index = index + 1
        return self.tempModel()[self.LineType[index]]

    @property
    def NodeArclengths(self):
        count = ctypes.c_int()
        status = ctypes.c_int()
        _lib.C_GetNodeArclengths(
            self.handle,
            None,
            ctypes.byref(count),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        arclengths = (ctypes.c_double * count.value)()
        _lib.C_GetNodeArclengths(
            self.handle,
            ctypes.byref(arclengths),
            None,
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return array(arclengths)

    def checkArclengthRange(self, arclengthRange):
        if not isinstance(arclengthRange, ArclengthRange):
            raise TypeError(
                'arclengthRange parameter must None, or be an instance of class ArclengthRange')
        return ctypes.byref(arclengthRange)

    def RangeGraphXaxis(self, varName, arclengthRange=None):
        varID = ctypes.c_int(self.varID(varName))
        status = ctypes.c_int()
        if arclengthRange is None:  # support 9.1 and earlier
            numberOfPoints = _lib.C_GetRangeGraphNumOfPoints(
                self.handle,
                varID,
                ctypes.byref(status)
            )
        else:
            numberOfPoints = _lib.C_GetRangeGraphNumOfPoints2(
                self.handle,
                self.checkArclengthRange(arclengthRange),
                varID,
                ctypes.byref(status)
            )
        _CheckStatus(status)
        x = (ctypes.c_double * numberOfPoints)()

        model = self.tempModel()
        if model.state == ModelState.InStaticState or model.isFrequencyDomainDynamics:
            period = Period(pnStaticState)
        else:
            simStart = -model.general.StageDuration[0]
            period = SpecifiedPeriod(simStart, simStart)
        if arclengthRange is None:  # support 9.1 and earlier
            _fn('C_GetRangeGraph3')(
                self.handle,
                self.checkObjectExtra(oeLine()),
                self.checkPeriod(period),
                varID,
                ctypes.byref(x),
                None,
                None,
                None,
                None,
                None,
                None,
                ctypes.byref(status)
            )
        else:
            _fn('C_GetRangeGraph4')(
                self.handle,
                self.checkObjectExtra(oeLine()),
                self.checkPeriod(period),
                self.checkArclengthRange(arclengthRange),
                varID,
                ctypes.byref(x),
                None,
                None,
                None,
                None,
                None,
                None,
                ctypes.byref(status)
            )
        _CheckStatus(status)
        return array(x)

    def RangeGraph(self, varName, period=None, objectExtra=None, arclengthRange=None, stormDurationHours=None):

        def alloc():
            return (ctypes.c_double * numberOfPoints)()

        def allocIfCurveAvailable(curveName):
            if getattr(curveNames, curveName):
                return alloc()
            else:
                return None

        def arrayPtr(x):
            if x is None:
                return None
            else:
                return ctypes.byref(x)

        def returnArray(x):
            if x is None:
                return None
            else:
                return array(x)

        def frequencyDomainRangeGraph(period, objectExtra, arclengthRange, stormDurationHours):
            def objectExtraAtArcLength(X):
                if objectExtra is None:
                    result = oeArcLength(X)
                else:
                    result = ObjectExtra()
                    # clone objectExtra, see http://stackoverflow.com/q/1470343/
                    ctypes.pointer(result)[0] = objectExtra
                    result.LinePoint = ptArcLength
                    result.ArcLength = X
                return result

            if period.PeriodNum != pnWholeSimulation:
                raise DLLError('Invalid period for frequency domain range graph, must be pnWholeSimulation')
            staticStateRangeGraph = self.RangeGraph(varName, pnStaticState, objectExtra, arclengthRange)
            fdr = [self.FrequencyDomainResults(varName, objectExtraAtArcLength(X)) for X in staticStateRangeGraph.X]
            resultDict = {
                'X': staticStateRangeGraph.X,
                'StaticValue': staticStateRangeGraph.Mean,
                'StdDev': array(map(lambda fdr: fdr.StdDev, fdr)),
                'Upper': staticStateRangeGraph.Upper,
                'Lower': staticStateRangeGraph.Lower
            }
            if stormDurationHours is not None:
                resultDict['MPM'] = array(
                    map(lambda fdr: FrequencyDomainMPM(stormDurationHours * 3600.0, fdr.StdDev, fdr.Tz), fdr))
            return objectFromDict(resultDict)

        model = self.tempModel()
        period = self.createPeriodParameter(period)
        if model.isFrequencyDomainDynamics and period.PeriodNum != pnStaticState:
            return frequencyDomainRangeGraph(period, objectExtra, arclengthRange, stormDurationHours)

        varID = ctypes.c_int(self.varID(varName))
        curveNames = RangeGraphCurveNames()
        status = ctypes.c_int()
        _fn('C_GetRangeGraphCurveNames')(
            self.handle,
            self.checkObjectExtra(objectExtra),
            self.checkPeriod(period),
            varID,
            ctypes.byref(curveNames),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        if arclengthRange is None:  # support 9.1 and earlier
            numberOfPoints = _lib.C_GetRangeGraphNumOfPoints(
                self.handle,
                varID,
                ctypes.byref(status)
            )
        else:
            numberOfPoints = _lib.C_GetRangeGraphNumOfPoints2(
                self.handle,
                self.checkArclengthRange(arclengthRange),
                varID,
                ctypes.byref(status)
            )
        _CheckStatus(status)
        x = alloc()
        min = allocIfCurveAvailable('Min')
        max = allocIfCurveAvailable('Max')
        mean = allocIfCurveAvailable('Mean')
        stddev = allocIfCurveAvailable('StdDev')
        upper = allocIfCurveAvailable('Upper')
        lower = allocIfCurveAvailable('Lower')
        if arclengthRange is None:  # support 9.1 and earlier
            _fn('C_GetRangeGraph3')(
                self.handle,
                self.checkObjectExtra(objectExtra),
                self.checkPeriod(period),
                varID,
                arrayPtr(x),
                arrayPtr(min),
                arrayPtr(max),
                arrayPtr(mean),
                arrayPtr(stddev),
                arrayPtr(upper),
                arrayPtr(lower),
                ctypes.byref(status)
            )
        else:
            _fn('C_GetRangeGraph4')(
                self.handle,
                self.checkObjectExtra(objectExtra),
                self.checkPeriod(period),
                self.checkArclengthRange(arclengthRange),
                varID,
                arrayPtr(x),
                arrayPtr(min),
                arrayPtr(max),
                arrayPtr(mean),
                arrayPtr(stddev),
                arrayPtr(upper),
                arrayPtr(lower),
                ctypes.byref(status)
            )
        _CheckStatus(status)

        return objectFromDict({
            'X': returnArray(x),
            'Min': returnArray(min),
            'Max': returnArray(max),
            'Mean': returnArray(mean),
            'StdDev': returnArray(stddev),
            'Upper': returnArray(upper),
            'Lower': returnArray(lower)
        })

    def _SaveExternalFile(self, filename, filetype, params=None):
        filename = _PrepareString(filename)
        status = ctypes.c_int()
        _fn('C_SaveExternalProgramFile')(
            self.handle,
            filetype,
            ctypes.byref(params) if params else None,
            _pchar(filename),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def SaveShear7datFile(self, filename):
        self._SaveExternalFile(filename, eftShear7dat)

    def SaveShear7mdsFile(self, filename, firstMode=-1, lastMode=-1):
        params = Shear7MdsFileParameters()
        params.FirstMode = firstMode
        params.LastMode = lastMode
        self._SaveExternalFile(filename, eftShear7mds, params)

    def SaveShear7outFile(self, filename):
        self._SaveExternalFile(filename, eftShear7out)

    def SaveShear7pltFile(self, filename):
        self._SaveExternalFile(filename, eftShear7plt)

    def SaveShear7anmFile(self, filename):
        self._SaveExternalFile(filename, eftShear7anm)

    def SaveShear7dmgFile(self, filename):
        self._SaveExternalFile(filename, eftShear7dmg)

    def SaveShear7fatFile(self, filename):
        self._SaveExternalFile(filename, eftShear7fat)

    def SaveShear7strFile(self, filename):
        self._SaveExternalFile(filename, eftShear7str)

    def SaveShear7out1File(self, filename):
        self._SaveExternalFile(filename, eftShear7out1)

    def SaveShear7out2File(self, filename):
        self._SaveExternalFile(filename, eftShear7out2)

    def SaveShear7allOutputFiles(self, basename):
        self._SaveExternalFile(basename, eftShear7allOutput)

    def SaveVIVAInputFiles(self, dirname):
        self._SaveExternalFile(dirname, eftVIVAInput)

    def SaveVIVAOutputFiles(self, basename):
        self._SaveExternalFile(basename, eftVIVAOutput)

    def SaveVIVAModesFiles(self, dirname, firstMode=-1, lastMode=-1):
        params = VIVAModesFilesParameters()
        params.FirstMode = firstMode
        params.LastMode = lastMode
        self._SaveExternalFile(dirname, eftVIVAModes, params)

    def SaveLineClashingReport(self, filename, period=None):
        parameters = LineClashingReportParameters()
        parameters.Period = self.createPeriodParameter(period)
        _SaveSpreadsheet(self.handle, sptLineClashingReport, filename, parameters)


class OrcaFlexVesselObject(OrcaFlexObject):

    def SaveDisplacementRAOsSpreadsheet(self, filename):
        _SaveSpreadsheet(self.handle, sptVesselDisplacementRAOs, filename, None)

    def SaveSpectralResponseSpreadsheet(self, filename):
        _SaveSpreadsheet(self.handle, sptVesselSpectralResponse, filename, None)


class OrcaFlexWizardObject(OrcaFlexObject):

    def InvokeWizard(self):
        status = ctypes.c_int()
        _lib.C_InvokeWizard(
            self.handle,
            ctypes.byref(status)
        )
        _CheckStatus(status)


class FatigueAnalysis(DataObject):

    def __init__(self, filename=None):
        handle = Handle()
        status = ctypes.c_int()
        _lib.C_CreateFatigue(ctypes.byref(handle), ctypes.byref(status))
        _CheckStatus(status)
        DataObject.__init__(self, handle)
        self.progressHandler = None
        if filename:
            self.Load(filename)

    def __del__(self):
        try:
            status = ctypes.c_int()
            _lib.C_DestroyFatigue(self.handle, ctypes.byref(status))
            # no point checking status here since we can't really do anything about a failure
        except:
            # swallow this since we get exceptions when Python terminates unexpectedly (e.g. CTRL+Z)
            pass

    def isBuiltInAttribute(self, name):
        return name == 'progressHandler' or DataObject.isBuiltInAttribute(self, name)

    def Load(self, filename):
        _doFileOperation(self.handle, 'LoadFatigue', filename)

    def Save(self, filename):
        _doFileOperation(self.handle, 'SaveFatigue', filename)

    def Calculate(self, resultsFilename=None):
        if self.progressHandler:
            def callback(handle, progress, cancel):
                cancel[0] = _ProgressCallbackCancel(self.progressHandler(self, progress))
            progress = ctypes.WINFUNCTYPE(None, Handle, _pchar, ctypes.POINTER(BOOL))(callback)
        else:
            progress = None
        resultsFilename = _PrepareString(resultsFilename)
        status = ctypes.c_int()
        _fn('C_CalculateFatigue')(
            self.handle,
            _pchar(resultsFilename),
            progress,
            ctypes.byref(status)
        )
        _CheckStatus(status)


class LinkedStatistics(object):

    def __init__(self, object, varNames, period=None, objectExtra=None):
        self.handle = Handle()
        self.object = object

        varCount = len(varNames)
        vars = (ctypes.c_int * varCount)()
        for index in range(varCount):
            vars[index] = ctypes.c_int(object.varID(varNames[index]))

        status = ctypes.c_int()
        _fn('C_OpenLinkedStatistics2')(
            self.object.handle,
            self.object.checkObjectExtra(objectExtra),
            self.object.checkPeriod(period),
            ctypes.c_int(varCount),
            ctypes.byref(vars),
            ctypes.byref(self.handle),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def __del__(self):
        try:
            status = ctypes.c_int()
            _lib.C_CloseLinkedStatistics(self.handle, ctypes.byref(status))
            # no point checking status here since we can't really do anything about a failure
        except:
            # swallow this since we get exceptions when Python terminates unexpectedly (e.g. CTRL+Z)
            pass

    def Query(self, varName, linkedVarName):
        result = StatisticsQuery()
        status = ctypes.c_int()
        _lib.C_QueryLinkedStatistics(
            self.handle,
            ctypes.c_int(self.object.varID(varName)),
            ctypes.c_int(self.object.varID(linkedVarName)),
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.asObject()

    def TimeSeriesStatistics(self, varName):
        result = TimeSeriesStats()
        status = ctypes.c_int()
        _lib.C_CalculateLinkedStatisticsTimeSeriesStatistics(
            self.handle,
            ctypes.c_int(self.object.varID(varName)),
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.asObject()


class ExtremeStatistics(object):

    def __init__(self, values, sampleInterval):
        self.handle = Handle()
        values = (ctypes.c_double * len(values))(*values)
        status = ctypes.c_int()
        _lib.C_OpenExtremeStatistics(
            len(values),
            ctypes.byref(values),
            ctypes.c_double(sampleInterval),
            ctypes.byref(self.handle),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def __del__(self):
        try:
            status = ctypes.c_int()
            _lib.C_CloseExtremeStatistics(self.handle, ctypes.byref(status))
            # no point checking status here since we can't really do anything about a failure
        except:
            # swallow this since we get exceptions when Python terminates unexpectedly (e.g. CTRL+Z)
            pass

    def ExcessesOverThresholdCount(self, specification=None):
        if specification is not None:
            if not isinstance(specification, ExtremeStatisticsSpecification):
                raise TypeError(
                    'specification must be an instance of ExtremeStatisticsSpecification')
            specification = ctypes.byref(specification)
        status = ctypes.c_int()
        count = _lib.C_CalculateExtremeStatisticsExcessesOverThreshold(
            self.handle,
            specification,
            None,
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return count

    def ExcessesOverThreshold(self, specification=None):
        count = self.ExcessesOverThresholdCount(specification)
        values = (ctypes.c_double * count)()
        if specification is not None:
            specification = ctypes.byref(specification)
        status = ctypes.c_int()
        count = _lib.C_CalculateExtremeStatisticsExcessesOverThreshold(
            self.handle,
            specification,
            ctypes.byref(values),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return array(values)

    def Fit(self, specification):
        if not isinstance(specification, ExtremeStatisticsSpecification):
            raise TypeError('specification must be an instance of ExtremeStatisticsSpecification')
        status = ctypes.c_int()
        _lib.C_FitExtremeStatistics(
            self.handle,
            ctypes.byref(specification),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def Query(self, query):
        if not isinstance(query, ExtremeStatisticsQuery):
            raise TypeError('query must be an instance of ExtremeStatisticsQuery')
        result = ExtremeStatisticsOutput()
        status = ctypes.c_int()
        _lib.C_QueryExtremeStatistics(
            self.handle,
            ctypes.byref(query),
            ctypes.byref(result),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.asObject()

    def ToleranceIntervals(self, SimulatedDataSetCount=1000):

        def intervalGenerator():
            intervals = (Interval * self.ExcessesOverThresholdCount())()
            status = ctypes.c_int()
            _lib.C_SimulateToleranceIntervals(
                self.handle,
                SimulatedDataSetCount,
                ctypes.byref(intervals),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            for interval in intervals:
                yield (interval.Lower, interval.Upper)

        return tuple(intervalGenerator())


class Modes(object):

    def __init__(self, obj, specification=None):
        self.handle = Handle()

        if specification is None:
            specification = ModalAnalysisSpecification()
        self.shapesAvailable = bool(specification.CalculateShapes)
        self.isWholeSystem = isinstance(obj, Model)

        dofCount = ctypes.c_int()
        modeCount = ctypes.c_int()
        status = ctypes.c_int()
        _lib.C_CreateModes(
            obj.handle,
            ctypes.byref(specification),
            ctypes.byref(self.handle),
            ctypes.byref(dofCount),
            ctypes.byref(modeCount),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        self.dofCount = dofCount.value
        self.modeCount = modeCount.value

        nodeNumber = (ctypes.c_int * self.dofCount)()
        dof = (ctypes.c_int * self.dofCount)()
        _lib.C_GetModeDegreeOfFreedomDetails(
            self.handle,
            ctypes.byref(nodeNumber),
            ctypes.byref(dof),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        self.nodeNumber = tuple(nodeNumber)
        self.dof = tuple(dof)

        def ownerGenerator():
            ownerHandle = (Handle * self.dofCount)()
            status = ctypes.c_int()
            _lib.C_GetModeDegreeOfFreedomOwners(
                self.handle,
                ctypes.byref(ownerHandle),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            latestOwner = None
            if self.isWholeSystem:
                model = obj
            else:
                model = obj.tempModel()
            for handle in ownerHandle:
                if (latestOwner is None) or (latestOwner.handle.value != handle):
                    latestOwner = model.createOrcaFlexObject(Handle(handle))
                yield latestOwner
        self.owner = tuple(ownerGenerator())

        modeNumber = (ctypes.c_int * self.modeCount)()
        period = (ctypes.c_double * self.modeCount)()
        _lib.C_GetModeSummary(
            self.handle,
            ctypes.byref(modeNumber),
            ctypes.byref(period),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        self.modeNumber = tuple(modeNumber)
        self.period = array(period)

    def __del__(self):
        try:
            status = ctypes.c_int()
            _lib.C_DestroyModes(self.handle, ctypes.byref(status))
            # no point checking status here since we can't really do anything about a failure
        except:
            # swallow this since we get exceptions when Python terminates unexpectedly (e.g. CTRL+Z)
            pass

    def modeDetails(self, index):
        if self.shapesAvailable:
            details = ModeDetails(self.dofCount)
        else:
            details = ModeDetails()
        status = ctypes.c_int()
        _lib.C_GetModeDetails(
            self.handle,
            ctypes.c_int(index),
            ctypes.byref(details),
            ctypes.byref(status)
        )
        _CheckStatus(status)
        if self.shapesAvailable:
            ShapeWrtGlobal = array(details.ShapeWrtGlobal)
            ShapeWrtLocal = array(details.ShapeWrtLocal)
        else:
            ShapeWrtGlobal = None
            ShapeWrtLocal = None
        resultDict = {
            'modeNumber': details.ModeNumber,
            'period': details.Period,
            'shapeWrtGlobal': ShapeWrtGlobal,
            'shapeWrtLocal': ShapeWrtLocal
        }
        if hasattr(details, 'ModeType'):
            resultDict['modeType'] = details.ModeType
            resultDict['percentageInInlineDirection'] = details.PercentageInInlineDirection
            resultDict['percentageInAxialDirection'] = details.PercentageInAxialDirection
            resultDict['percentageInTransverseDirection'] = details.PercentageInTransverseDirection
            resultDict['percentageInRotationalDirection'] = details.PercentageInRotationalDirection
        else:
            resultDict['modeType'] = mtNotAvailable
            resultDict['percentageInInlineDirection'] = OrcinaNullReal()
            resultDict['percentageInAxialDirection'] = OrcinaNullReal()
            resultDict['percentageInTransverseDirection'] = OrcinaNullReal()
            resultDict['percentageInRotationalDirection'] = OrcinaNullReal()
        return objectFromDict(resultDict)

    def _modeLoadOutputPointCount(self):
        result = ctypes.c_int()
        status = ctypes.c_int()
        _lib.C_GetModeLoadOutputPoints(
            self.handle,
            ctypes.byref(result),
            None,
            ctypes.byref(status)
        )
        _CheckStatus(status)
        return result.value

    @property
    def modeLoadOutputPoints(self):

        def pointGenerator():

            class ModeLoadOutputPoint(PackedStructure):

                _fields_ = [
                    ('Owner', Handle),
                    ('Arclength', ctypes.c_double)
                ]

            outputPoints = (ModeLoadOutputPoint * self._modeLoadOutputPointCount())()
            status = ctypes.c_int()
            _lib.C_GetModeLoadOutputPoints(
                self.handle,
                None,
                ctypes.byref(outputPoints),
                ctypes.byref(status)
            )
            _CheckStatus(status)

            latestOwner = None
            model = None
            for point in outputPoints:
                if model is None:
                    modelHandle = Handle()
                    _lib.C_GetModelHandle(
                        point.Owner,
                        ctypes.byref(modelHandle),
                        ctypes.byref(status)
                    )
                    _CheckStatus(status)
                    model = Model(handle=modelHandle)
                if (latestOwner is None) or (latestOwner.handle.value != point.Owner):
                    latestOwner = model.createOrcaFlexObject(Handle(point.Owner))
                yield objectFromDict({
                    'owner': latestOwner,
                    'arclength': point.Arclength
                })

        return tuple(pointGenerator())

    def modeLoad(self, index):

        def loadGenerator():

            class ModeLoad(PackedStructure):

                _fields_ = [
                    ('Force', ctypes.c_double * 3),
                    ('Moment', ctypes.c_double * 3)
                ]

            loads = (ModeLoad * self._modeLoadOutputPointCount())()
            status = ctypes.c_int()
            _lib.C_GetModeLoad(
                self.handle,
                ctypes.c_int(index),
                ctypes.byref(loads),
                ctypes.byref(status)
            )
            _CheckStatus(status)

            for load in loads:
                yield objectFromDict({
                    'force': array(load.Force),
                    'moment': array(load.Moment)
                })

        return tuple(loadGenerator())


class AVIFile(object):

    def __init__(self, filename, codec, interval, width, height):
        self.width = width
        self.height = height
        self.handle = Handle()
        filename = _PrepareString(filename)
        status = ctypes.c_int()
        _fn('C_AVIFileInitialise')(
            ctypes.byref(self.handle),
            _pchar(filename),
            ctypes.byref(AVIFileParameters(codec, interval)),
            ctypes.byref(status)
        )
        _CheckStatus(status)

    def AddFrame(self, model, drawTime, viewParameters=None):
        if viewParameters is None:
            viewParameters = model.defaultViewParameters
        else:
            model.checkViewParameters(viewParameters)
        viewParameters.Height = self.height
        viewParameters.Width = self.width
        frame = HBITMAP()
        saveDrawTime = model.simulationDrawTime
        try:
            model.simulationDrawTime = drawTime
            status = ctypes.c_int()
            _lib.C_CreateModel3DViewBitmap(
                model.handle,
                ctypes.byref(viewParameters),
                ctypes.byref(frame),
                ctypes.byref(status)
            )
            _CheckStatus(status)
            try:
                _lib.C_AVIFileAddBitmap(
                    self.handle,
                    frame,
                    ctypes.byref(status)
                )
                _CheckStatus(status)
            finally:
                if ctypes.windll.gdi32.DeleteObject(frame) == 0:
                    raise ctypes.WinError()
        finally:
            model.simulationDrawTime = saveDrawTime

    def Close(self):
        if self.handle != 0:
            status = ctypes.c_int()
            _lib.C_AVIFileFinalise(
                self.handle,
                ctypes.byref(status)
            )
            _CheckStatus(status)
            self.handle = 0


def TimeSeriesStatistics(values, SampleInterval):
    values = (ctypes.c_double * len(values))(*values)
    result = TimeSeriesStats()
    status = ctypes.c_int()
    _lib.C_CalculateTimeSeriesStatistics(
        ctypes.byref(values),
        len(values),
        ctypes.c_double(SampleInterval),
        ctypes.byref(result),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return result.asObject()


def AnalyseExtrema(values):
    values = (ctypes.c_double * len(values))(*values)
    Max = ctypes.c_double()
    Min = ctypes.c_double()
    IndexOfMax = ctypes.c_int()
    IndexOfMin = ctypes.c_int()
    status = ctypes.c_int()
    _lib.C_AnalyseExtrema(
        ctypes.byref(values),
        len(values),
        ctypes.byref(Max),
        ctypes.byref(Min),
        ctypes.byref(IndexOfMax),
        ctypes.byref(IndexOfMin),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return objectFromDict({
        'Min': Min.value,
        'Max': Max.value,
        'IndexOfMin': IndexOfMin.value,
        'IndexOfMax': IndexOfMax.value
    })


def TimeHistorySummary(TimeHistorySummaryType, Times, Values, spectralDensityFundamentalFrequency=None):
    if TimeHistorySummaryType == thstSpectralDensity:
        if len(Times) != len(Values):
            raise ValueError('Times and Values must have the same length')
        times = ctypes.byref((ctypes.c_double * len(Times))(*Times))
    else:
        times = 0
    values = (ctypes.c_double * len(Values))(*Values)
    handle = Handle()
    summaryValueCount = ctypes.c_int()
    status = ctypes.c_int()
    if spectralDensityFundamentalFrequency is None:
        _lib.C_CreateTimeHistorySummary(
            TimeHistorySummaryType,
            len(Values),
            times,
            ctypes.byref(values),
            ctypes.byref(handle),
            ctypes.byref(summaryValueCount),
            ctypes.byref(status)
        )
    else:
        specification = TimeHistorySummarySpecification()
        specification.SpectralDensityFundamentalFrequency = spectralDensityFundamentalFrequency
        _lib.C_CreateTimeHistorySummary2(
            TimeHistorySummaryType,
            len(Values),
            ctypes.byref(specification),
            times,
            ctypes.byref(values),
            ctypes.byref(handle),
            ctypes.byref(summaryValueCount),
            ctypes.byref(status)
        )
    _CheckStatus(status)
    try:
        x = (ctypes.c_double * summaryValueCount.value)()
        y = (ctypes.c_double * summaryValueCount.value)()
        _lib.C_GetTimeHistorySummaryValues(
            handle,
            ctypes.byref(x),
            ctypes.byref(y),
            ctypes.byref(status)
        )
        _CheckStatus(status)
    finally:
        _lib.C_DestroyTimeHistorySummary(
            handle,
            ctypes.byref(status)
        )
        _CheckStatus(status)
    return GraphCurve(x, y)


def SpectralDensity(times, values, fundamentalFrequency=None):
    return TimeHistorySummary(thstSpectralDensity, times, values, fundamentalFrequency)


def EmpiricalDistribution(values):
    return TimeHistorySummary(thstEmpiricalDistribution, None, values)


def RainflowHalfCycles(values):
    return TimeHistorySummary(thstRainflowHalfCycles, None, values).X


def FrequencyDomainMPM(stormDuration, StdDev, Tz):
    result = ctypes.c_double()
    status = ctypes.c_int()
    _lib.C_GetFrequencyDomainMPM(
        ctypes.c_double(stormDuration),
        ctypes.c_double(StdDev),
        ctypes.c_double(Tz),
        ctypes.byref(result),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return result.value


def MoveObjectDisplacementSpecification(displacement):
    result = MoveObjectSpecification()
    result.MoveSpecifiedBy = sbDisplacement
    result.Displacement = displacement
    return result


def MoveObjectPolarDisplacementSpecification(direction, distance):
    result = MoveObjectSpecification()
    result.MoveSpecifiedBy = sbPolarDisplacement
    result.PolarDisplacementDirection = direction
    result.PolarDisplacementDistance = distance
    return result


def MoveObjectNewPositionSpecification(referenceObject, referencePointIndex, newPosition):
    result = MoveObjectSpecification()
    result.MoveSpecifiedBy = sbNewPosition
    result.NewPositionReferencePoint = MoveObjectPoint(referenceObject, referencePointIndex)
    result.NewPosition = newPosition
    return result


def MoveObjectRotationSpecification(angle, centre):
    result = MoveObjectSpecification()
    result.MoveSpecifiedBy = sbRotation
    result.RotationAngle = angle
    result.RotationCentre = centre
    return result


def MoveObjects(specification, points):
    points = (MoveObjectPoint * len(points))(*points)
    status = ctypes.c_int()
    _lib.C_MoveObjects(ctypes.byref(specification), len(points), ctypes.byref(points), ctypes.byref(status))
    _CheckStatus(status)


def ExchangeObjects(object1, object2):
    status = ctypes.c_int()
    _lib.C_ExchangeObjects(object1.handle, object2.handle, ctypes.byref(status))
    _CheckStatus(status)


def SortObjects(objects, key):
    objects = list(objects)
    N = len(objects)
    indices = list(range(N))
    locations = indices[:]
    order = indices[:]
    order.sort(key=lambda i: key(objects[i]))
    indices = list(range(N))
    locations = indices[:]
    for i in range(N):
        index1, index2 = i, locations[order[i]]
        ExchangeObjects(objects[indices[index1]], objects[indices[index2]])
        locations[indices[index1]], locations[indices[index2]] = index2, index1
        indices[index1], indices[index2] = indices[index2], indices[index1]


def CalculateMooringStiffness(vessels):
    _requiresNumpy()
    status = ctypes.c_int()
    N = len(vessels)
    vesselHandles = (Handle * N)(*map(lambda vessel: vessel.handle, vessels))
    result = numpy.zeros(6 * N * 6 * N, dtype=numpy.float64)
    _lib.C_CalculateMooringStiffness(len(vesselHandles), vesselHandles, result, status)
    _CheckStatus(status)
    return result.reshape((6 * N, 6 * N))


def DLLVersion():
    v = (_char * 16)()
    status = ctypes.c_int()
    _fn('C_GetDLLVersion')(
        None,  # RequiredVersion
        ctypes.byref(v),
        ctypes.byref(ctypes.c_int()),  # OK
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return _DecodeString(v.value)


def DLLLocation():
    fileName = (_char * MAX_PATH)()
    func = ctypes.windll.kernel32['GetModuleFileNameW' if _UseUnicodeOrcFxAPI else 'GetModuleFileNameA']
    returnValue = func(
        _lib._handle,
        ctypes.byref(fileName),
        MAX_PATH
    )
    if returnValue == 0:
        raise ctypes.WinError()
    return _DecodeString(fileName.value)


def BinaryFileType(filename):
    filename = _PrepareString(filename)
    ft = ctypes.c_int()
    status = ctypes.c_int()
    _fn('C_GetBinaryFileType')(
        _pchar(filename),
        ctypes.byref(ft),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return FileType[ft.value]


def FileCreatorVersion(filename):
    filename = _PrepareString(filename)
    func = _fn('C_GetFileCreatorVersion')
    status = ctypes.c_int()
    length = func(
        _pchar(filename),
        None,
        ctypes.byref(status)
    )
    _CheckStatus(status)
    result = (_char * length)()
    func(
        filename,
        ctypes.byref(result),
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return _DecodeString(result.value)


def SolveEquation(calcY, initialX, targetY, params=None):
    if params:
        params = ctypes.byref(params)
    else:
        params = None

    def callback(Data, X, CallBackStatus):
        return calcY(X)
    SolveEquationCalcYProc = ctypes.WINFUNCTYPE(
        ctypes.c_double, ctypes.c_int, ctypes.c_double, ctypes.POINTER(ctypes.c_int))(callback)
    result = ctypes.c_double(initialX)
    status = ctypes.c_int()
    _lib.C_SolveEquation(
        ctypes.c_int(0),  # Data, but we don't use this mechanism since we can use closures instead
        SolveEquationCalcYProc,
        ctypes.byref(result),
        ctypes.c_double(targetY),
        params,
        ctypes.byref(status)
    )
    _CheckStatus(status)
    return result.value
