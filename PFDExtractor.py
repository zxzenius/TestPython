__author__ = 'zenius'

import win32com.client
import os.path


class HyExtractor:
    def __init__(self, HyFile):
        self.HyFile = HyFile

    def isValid(self):
        return os.path.isfile(self.HyFile)

    def Test(self):
        if not self.isValid():
            print('FileNameError!')
            return False
        self.hyApp = win32com.client.Dispatch('Hysys.Application')
        self.hyCase = self.hyApp.SimulationCases.Open(self.HyFile)
        self.hyStreams = self.hyCase.Flowsheet.MaterialStreams
        try:
            for hyStream in self.hyStreams:
                print(hyStream.ActualGasFlowValue)
                #print(self.ExtractNote(hyStream))
        finally:
            self.hyApp = None


    def ExtractNote(self, Stream):
        NoteData = []
        NoteData.append(Stream.ActualGasFlowValue)

        #        NoteData.append(Stream.ActualLiqFlowValue)

#        NoteData.append(Stream.ActualVolumeFlowValue)
#        NoteData.append(Stream.AvgLiqDensityValue)
#        NoteData.append(Stream.BOBubblePointPressureValue)
#        NoteData.append(Stream.BOBubblePointTemperatureValue)
#        NoteData.append(Stream.BOGasOilRatioValue)
#        NoteData.append(Stream.BOMassEnthalpyValue)
#        NoteData.append(Stream.BOMassFlowValue)
#        NoteData.append(Stream.BOOilFormationVolumeFactorValue)
#        NoteData.append(Stream.BOOilViscosityValue)
#        NoteData.append(Stream.BOPressureValue)
#        NoteData.append(Stream.BOSolutionGORValue)
#        NoteData.append(Stream.BOSpecificGravityValue)
#        NoteData.append(Stream.BOSurfaceTensionValue)
#        NoteData.append(Stream.BOTemperatureInVMValue)
#        NoteData.append(Stream.BOTemperatureValue)
#        NoteData.append(Stream.BOViscosityCoefficientAValue)
#        NoteData.append(Stream.BOViscosityCoefficientBValue)
#        NoteData.append(Stream.BOViscosityValue)
#        NoteData.append(Stream.BOVolumetricFlowValue)
#        NoteData.append(Stream.BOWaterCutValue)
#        NoteData.append(Stream.BOWatsonKValue)
#        NoteData.append(Stream.ComponentMassFlowValue)
#        NoteData.append(Stream.ComponentMolarFlowValue)
#        NoteData.append(Stream.ComponentMolarFractionValue)
#        NoteData.append(Stream.ComponentVolumeFlowValue)
#        NoteData.append(Stream.ComponentVolumeFractionValue)
#        NoteData.append(Stream.CompressibilityValue)
#        NoteData.append(Stream.CpCvValue)
#        NoteData.append(Stream.EnthalpyEstimateValue)
#        NoteData.append(Stream.FlowEstimateValue)
#        NoteData.append(Stream.HeatFlowValue)
#        NoteData.append(Stream.HeatOfVapValue)
#        NoteData.append(Stream.HeavyLiquidFractionValue)
#        NoteData.append(Stream.HigherHeatValueValue)
#        NoteData.append(Stream.IdealLiquidVolumeFlowValue)
#        NoteData.append(Stream.IsEnergyStream)
#        NoteData.append(Stream.IsValid)
#        NoteData.append(Stream.KineticViscosityValue)
#        NoteData.append(Stream.LightLiquidFractionValue)
#        NoteData.append(Stream.LiquidFractionValue)
#        NoteData.append(Stream.LowerHeatValueValue)
#        NoteData.append(Stream.MassDensityValue)
#        NoteData.append(Stream.MassEnthalpyValue)
#        NoteData.append(Stream.MassEntropyValue)
#        NoteData.append(Stream.MassFlowValue)
#        NoteData.append(Stream.MassHeatCapacityValue)
#        NoteData.append(Stream.MassHeatOfVapValue)
#        NoteData.append(Stream.MassHigherHeatValueValue)
#        NoteData.append(Stream.MassLowerHeatValueValue)
#        NoteData.append(Stream.MolarDensityValue)
#        NoteData.append(Stream.MolarEnthalpyValue)
#        NoteData.append(Stream.MolarEntropyValue)
#        NoteData.append(Stream.MolarFlowValue)
#        NoteData.append(Stream.MolarHeatCapacityValue)
#        NoteData.append(Stream.MolarVolumeValue)
#        NoteData.append(Stream.MolecularWeightValue)
#        NoteData.append(Stream.Name)
#        NoteData.append(Stream.PowerValue)
#        NoteData.append(Stream.PressureCO2Value)
#        NoteData.append(Stream.PressureValue)
#        NoteData.append(Stream.SGAirValue)
#        NoteData.append(Stream.StdGasFlowValue)
#        NoteData.append(Stream.StdLiqMassDensityValue)
#        NoteData.append(Stream.StdLiqVolFlowValue)
#        NoteData.append(Stream.StreamDescription)
#        NoteData.append(Stream.SurfaceTensionValue)
#        NoteData.append(Stream.TaggedName)
#        NoteData.append(Stream.TemperatureEstimateValue)
#        NoteData.append(Stream.TemperatureValue)
#        NoteData.append(Stream.ThermalConductivityValue)
#        NoteData.append(Stream.TypeName)
#        NoteData.append(Stream.UniqueID)
#        NoteData.append(Stream.VapourFractionValue)
#        NoteData.append(Stream.ViscosityValue)
#        NoteData.append(Stream.VisibleTypeName)
#        NoteData.append(Stream.WatsonValue)



if __name__ == '__main__':
    HyFile = 'd:\\temp\\test.hsc'
    ProtoType = HyExtractor(HyFile)
    try:
        ProtoType.Test()
    finally:
        pass