// =DOC FILE Interface for the CProFumeCalculator class
// =DOC TEXT Declares CProFumeCalculator class to perform all fumigation calculations
//////////////////////////////////////////////////////////////////

// ProFumeCalculator.h : Declaration of the CProFumeCalculator

#ifndef __PROFUMECALCULATOR_H_
#define __PROFUMECALCULATOR_H_

#include "resource.h"       // main symbols
#include "DoseData.h"


class DArea;
class DDoseCalc;
class DMonitorPoint;
class DValidateAreaMonPtInput;
class DValidateDoseInput;
class DValidateHalfLifeInput;
class DValidateConcentration;

#include "AreaMonPtData.h"
#include "ConcentrationData.h"


/////////////////////////////////////////////////////////////////////////////
// CProFumeCalculator
// =DOC SECTION Class CProFumeCalculator
// =DOC SECTION Class containing all ProFume calculator methods
// =DOC TEXT class CProFumeCalculator extends CComObjectRootEx<CComSingleThreadModel>, CComCoClass<CProFumeCalculator, &CLSID_ProFumeCalculator>, &IID_IProFumeCalculator, &LIBID_DPCALCLib>.
// =DOC TEXT Class containing all ProFume calculator methods
//////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE CProFumeCalculator : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CProFumeCalculator, &CLSID_ProFumeCalculator>,
	public ISupportErrorInfo,
	public IDispatchImpl<IProFumeCalculator, &IID_IProFumeCalculator, &LIBID_DPCALCLib>
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	CProFumeCalculator();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   ~CProFumeCalculator();

DECLARE_REGISTRY_RESOURCEID(IDR_PROFUMECALCULATOR)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CProFumeCalculator)
	COM_INTERFACE_ENTRY(IProFumeCalculator)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IProFumeCalculator
public:
   // =DOC SECTION Method GetAreaStatus( /*[in]*/ double TargetConcentrationTime, /*[out]*/ double* AchievedConcentrationTime, /*[out]*/ double* AchievedDose, /*[out]*/ double* Hlt, /*[out]*/ double* ProjectedConcentrationTime, /*[out]*/ double* ProjectedTime, /*[out]*/ double* MeanNTC, /*[out]*/ int* Status, /*[out]*/ double* AddProfume, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC SECTION Retrieve the status of an area's after adding all of the area monitor point status values (SetAreaMonitorPoint()).
   // =DOC TEXT public GetAreaStatus( /*[in]*/  double  TargetConcentrationTime, /*[out]*/ double* AchievedConcentrationTime, /*[out]*/ double* AchievedDose, /*[out]*/ double* Hlt, /*[out]*/ double* ProjectedConcentrationTime, /*[out]*/ double* ProjectedTime, /*[out]*/ double* MeanNTC, /*[out]*/ int* Status, /*[out]*/ double* AddProfume, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC TEXT /*[in]*/  double TargetConcentrationTime: Target CT to achieve
   // =DOC TEXT /*[out]*/ double* AchievedConcentrationTime: CT achieved at current time
   // =DOC TEXT /*[out]*/ double* AchievedDose: Dose achieved at current time
   // =DOC TEXT /*[out]*/ double* Hlt: Calculated half loss time for area
   // =DOC TEXT /*[out]*/ double* ProjectedConcentrationTime: Projected CT achieved
   // =DOC TEXT /*[out]*/ double* ProjectedTime: Projected total fumigation time
   // =DOC TEXT /*[out]*/ double* MeanNTC: Average new target concentration for the area
   // =DOC TEXT /*[out]*/ int* Status: Status mode of the area
   // =DOC TEXT /*[out]*/ double* AddProfume: Amount of Profume to add to area (kg)
   // =DOC TEXT /*[out]*/ BSTR* Msg: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Retrieve the status of an area after adding all of the area's monitor point status values (SetAreaMonitorPoint()).
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetAreaStatus)( /*[in]*/  double  TargetConcentrationTime,
                             /*[out]*/ double* AchievedConcentrationTime, 
                             /*[out]*/ double* AchievedDose, 
                             /*[out]*/ double* Hlt,
                             /*[out]*/ double* ProjectedConcentrationTime,
                             /*[out]*/ double* ProjectedTime,
                             /*[out]*/ double* MeanNTC,
                             /*[out]*/ int*    Status,
                             /*[out]*/ double* AddProfume,
                             /*[out]*/ BSTR*   Msg );
	
   // =DOC SECTION Method ClearAreaMonitorPoints(): STDMETHOD
   // =DOC SECTION Clear the list of area monitor point status values in m_padataAreaMonPt.
   // =DOC TEXT public ClearAreaMonitorPoints(): STDMETHOD
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Clear the list of area monitor point status values in m_padataAreaMonPt.
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(ClearAreaMonitorPoints)();

   // =DOC SECTION Method SetAreaMonitorPoint( /*[in]*/ double AchievedConcentrationTime, /*[in]*/ double AchievedDose, /*[in]*/ double DeltaDose, /*[in]*/ double ElapsedTime, /*[in]*/ double NewTargetConcentration, /*[in]*/ double Hlt, /*[in]*/ double ProjectedConcentrationTime, /*[in]*/ double ProjectedTime, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC SECTION Set the properties of an Area Monitor Point
   // =DOC TEXT public SetAreaMonitorPoint( /*[in]*/ double AchievedConcentrationTime, /*[in]*/ double AchievedDose, /*[in]*/ double DeltaDose, /*[in]*/ double ElapsedTime, /*[in]*/ double NewTargetConcentration, /*[in]*/ double Hlt, /*[in]*/ double ProjectedConcentrationTime, /*[in]*/ double ProjectedTime, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC TEXT /*[in]*/ double AchievedConcentrationTime: CT achieved at current time
   // =DOC TEXT /*[in]*/ double AchievedDose: Dose achieved at current time
   // =DOC TEXT /*[in]*/ double DeltaDose: Delta (change in) dose
   // =DOC TEXT /*[in]*/ double ElapsedTime: Elapsed hours since start of exposure
   // =DOC TEXT /*[in]*/ double NewTargetConcentration: Concentration target for one hour in the future
   // =DOC TEXT /*[in]*/ double Hlt: HLT Half-loss time
   // =DOC TEXT /*[in]*/ double ProjectedConcentrationTime: Projected CT achieved
   // =DOC TEXT /*[in]*/ double ProjectedTime: Projected time
   // =DOC TEXT /*[out]*/ BSTR* Msg: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Set the properties of an Area Monitor Point
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(SetAreaMonitorPoint)( /*[in]*/ double AchievedConcentrationTime, 
                                   /*[in]*/ double AchievedDose,
                                   /*[in]*/ double DeltaDose,
                                   /*[in]*/ double ElapsedTime,
                                   /*[in]*/ double NewTargetConcentration,
                                   /*[in]*/ double Hlt,
                                   /*[in]*/ double ProjectedConcentrationTime,
                                   /*[in]*/ double ProjectedTime,
                                   /*[out]*/ BSTR* Msg);
   
   
   // =DOC SECTION Method GetPt2PtHlt( /*[out]*/ double* LastPeakHLT, /*[out]*/ DATE* LastPeakDateTime, /*[out]*/ DATE* HLTDateTime, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC SECTION Get the point to point half loss time (HLT)
   // =DOC TEXT public  GetPt2PtHlt( /*[out]*/ double* LastPeakHLT, /*[out]*/ DATE* LastPeakDateTime, /*[out]*/ DATE* HLTDateTime, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC TEXT /*[out]*/ double* LastPeakHLT: The HLT at the point from the last peak (hrs)
   // =DOC TEXT /*[out]*/ DATE* LastPeakDateTime: The date/time of the start point (last peak) when determining HLT
   // =DOC TEXT /*[out]*/ DATE* HLTDateTime: The date/time of the end point when determining HLT
   // =DOC TEXT /*[out]*/ BSTR* Msg: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the point to point half loss time (HLT).  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetPt2PtHlt)( /*[out]*/ double* LastPeakHLT,
                           /*[out]*/ DATE* LastPeakDateTime,
                           /*[out]*/ DATE* HLTDateTime, 
                           /*[out]*/ BSTR* Msg );

   // =DOC SECTION Method GetVersion( /*[out, retval]*/ BSTR* Version ): STDMETHOD
   // =DOC SECTION Get the version number of the DAS ProFume calculator
   // =DOC TEXT public  GetVersion( /*[out, retval]*/ BSTR* Version ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ BSTR* Version: The version string
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the version number of the DAS ProFume calculator  
   // =DOC TEXT Return S_OK
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetVersion)( /*[out, retval]*/ BSTR* Version );

   // =DOC SECTION Method ClearConcentrations(): STDMETHOD
   // =DOC SECTION Clear the concentrations in m_padataConc
   // =DOC TEXT public  ClearConcentrations(): STDMETHOD
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Clear the concentrations in m_padataConc.  
   // =DOC TEXT Return S_OK
   //////////////////////////////////////////////////////////////////
	STDMETHOD(ClearConcentrations)();
	
   // =DOC SECTION Method SetConcentration( /*[in]*/ double Concentration, /*[in]*/ DATE Date, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC SECTION Add a new concentration value to m_padataConc
   // =DOC TEXT public  SetConcentration( /*[in]*/ double Concentration, /*[in]*/ DATE Date, /*[out]*/ BSTR* Msg ): STDMETHOD
   // =DOC TEXT /*[in]*/ double Concentration: The concentration value (g)
   // =DOC TEXT /*[in]*/ DATE Date: The date/time of the concentration value
   // =DOC TEXT /*[out]*/ BSTR* Msg: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Add a new concentration value to m_padataConc.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the add.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(SetConcentration)( /*[in]*/ double Concentration, 
                                /*[in]*/ DATE Date, 
                                /*[out]*/ BSTR* Msg );
	
   // =DOC SECTION Method GetMonitorPointStatus( /*[in]*/ double OriginalHLT, /*[out]*/ double* AchievedCT, /*[out]*/ double* AchievedDose, /*[out]*/ double* ActualOHLT, /*[out]*/ double* ProjectedCT, /*[out]*/ double* ProjectedFumigationTime, /*[out]*/ double* NewTargetConcentration, /*[out]*/ double* DeltaDose, /*[out]*/ double* ElapsedTime, /*[out]*/ DATE* LastPeakDateTime, /*[out]*/ DATE* HLTDateTime, /*[out]*/ BSTR* Msg): STDMETHOD
   // =DOC SECTION Retrieve the status of a monitor point after adding all of the monitor points time/concentration values (SetConcentration()).
   // =DOC TEXT public  GetMonitorPointStatus( /*[in]*/ double OriginalHLT, /*[out]*/ double* AchievedCT, /*[out]*/ double* AchievedDose, /*[out]*/ double* ActualOHLT, /*[out]*/ double* ProjectedCT, /*[out]*/ double* ProjectedFumigationTime, /*[out]*/ double* NewTargetConcentration, /*[out]*/ double* DeltaDose, /*[out]*/ double* ElapsedTime, /*[out]*/ DATE* LastPeakDateTime, /*[out]*/ DATE* HLTDateTime, /*[out]*/ BSTR* Msg): STDMETHOD
   // =DOC TEXT /*[in]*/ double OriginalHLT: The original HLT (hrs)
   // =DOC TEXT /*[out]*/ double* AchievedCT: The calculated achieved concentration
   // =DOC TEXT /*[out]*/ double* AchievedDose: The calculated achieved dose (g-hr/m^3)
   // =DOC TEXT /*[out]*/ double* ActualOHLT: The calculated HLT based on the last peak concentration and date/time (hrs)
   // =DOC TEXT /*[out]*/ double* ProjectedCT: The calculated projected CT achieved
   // =DOC TEXT /*[out]*/ double* ProjectedFumigationTime: The calculated total fumigation time required
   // =DOC TEXT /*[out]*/ double* NewTargetConcentration: The calculated concentration that we want to have in 1 hour
   // =DOC TEXT /*[out]*/ double* DeltaDose: The calculated difference between doses
   // =DOC TEXT /*[out]*/ double* ElapsedTime: The amount of time since exposure began (ms)
   // =DOC TEXT /*[out]*/ DATE* LastPeakDateTime: The date/time of the last peak concentration (starting HLT point)
   // =DOC TEXT /*[out]*/ DATE* HLTDateTime: The date/time of the last concentration during exposure (ending HLT point)
   // =DOC TEXT /*[out]*/ BSTR* Msg: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Retrieve the status of a monitor point after adding all of the monitor points time/concentration values (SetConcentration()).
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(GetMonitorPointStatus)( /*[in]*/ double OriginalHLT, 
                                     /*[out]*/ double* AchievedCT,
                                     /*[out]*/ double* AchievedDose,
                                     /*[out]*/ double* ActualOHLT,
                                     /*[out]*/ double* ProjectedCT,
                                     /*[out]*/ double* ProjectedFumigationTime,
                                     /*[out]*/ double* NewTargetConcentration,
                                     /*[out]*/ double* DeltaDose,
                                     /*[out]*/ double* ElapsedTime,
                                     /*[out]*/ DATE* LastPeakDateTime,
                                     /*[out]*/ DATE* HLTDateTime, 
                                     /*[out]*/ BSTR* Msg);

   // =DOC SECTION Method ClearDosePest(): STDMETHOD
   // =DOC SECTION Clear the pest data in m_astrPest
   // =DOC TEXT public  ClearDosePest(): STDMETHOD
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Clear the pest data in m_astrPest.  
   // =DOC TEXT Return S_OK
   //////////////////////////////////////////////////////////////////
	STDMETHOD(ClearDosePest)();

   // =DOC SECTION Method GetProfumeRequired( /*[out, retval]*/ double* ProfumeReq, /*[out]*/ double* Cylinders, /*[out]*/ double* Co, /*[out]*/ double* CT, /*[out]*/ double* Dose, /*[out]*/ BSTR* Error ): STDMETHOD
   // =DOC SECTION Get the amount of ProFume required for the fumigation.
   // =DOC TEXT public  GetProfumeRequired( /*[out, retval]*/ double* ProfumeReq, /*[out]*/ double* Cylinders, /*[out]*/ double* Co, /*[out]*/ double* CT, /*[out]*/ double* Dose, /*[out]*/ BSTR* Error ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ double* ProfumeReq: The calculated amount of ProFume required for the fumigation (kg)
   // =DOC TEXT /*[out]*/ double* Cylinders: The calculated number of cylinders required for the fumigation
   // =DOC TEXT /*[out]*/ double* Co: The calculated initial concentration (g/m^3)
   // =DOC TEXT /*[out]*/ double* CT: The calculated concentration time (g-hr/m^3)
   // =DOC TEXT /*[out]*/ double* Dose: The calculated dose (g-hr/m^3)
   // =DOC TEXT /*[out]*/ BSTR* Error: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the amount of ProFume required for the fumigation.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetProfumeRequired)( /*[out, retval]*/ double* ProfumeReq, 
                                  /*[out]*/ double* Cylinders, 
                                  /*[out]*/ double* Co, 
                                  /*[out]*/ double* CT, 
                                  /*[out]*/ double* Dose, 
                                  /*[out]*/ BSTR* Error );

   // =DOC SECTION Method SetDosePest( /*[in]*/ int PestIndex, /*[in]*/ int FumigationType, /*[in]*/ int LifeStage, /*[in]*/ int TreatmentType, /*[in]*/ int Commodity, /*[in]*/ double HLT, /*[in]*/ double Temp, /*[in]*/ double TimeExposure, /*[in]*/ double Volume, /*[out]*/ BSTR* Error ): STDMETHOD
   // =DOC SECTION Add a new pest for this dose calculation.
   // =DOC TEXT public  SetDosePest( /*[in]*/ int PestIndex, /*[in]*/ int FumigationType, /*[in]*/ int LifeStage, /*[in]*/ int TreatmentType, /*[in]*/ int Commodity, /*[in]*/ double HLT, /*[in]*/ double Temp, /*[in]*/ double TimeExposure, /*[in]*/ double Volume, /*[out]*/ BSTR* Error ): STDMETHOD
   // =DOC TEXT /*[in]*/ int PestIndex: Pest index
   // =DOC TEXT /*[in]*/ int FumigationType: Fumigation type index
   // =DOC TEXT /*[in]*/ int LifeStage: Life stage index
   // =DOC TEXT /*[in]*/ int TreatmentType: Treatment type index
   // =DOC TEXT /*[in]*/ int Commodity: Commodity index
   // =DOC TEXT /*[in]*/ double HLT: Estimated half loss time, HLT (hrs)
   // =DOC TEXT /*[in]*/ double Temp: Temperature of the fumigation area (C)
   // =DOC TEXT /*[in]*/ double TimeExposure: Time of exposure (hrs)
   // =DOC TEXT /*[in]*/ double Volume: Volume of the fumigation area (m^3)
   // =DOC TEXT /*[in]*/ double LaodFactor: Load Factor needed to use in the computation of HLT Sorption
   // =DOC TEXT /*[in]*/ int IsComplete: Will indicate if the criteria are complete
   // =DOC TEXT /*[in]*/ int ICounter: This will indicate 
   // =DOC TEXT /*[out]*/ BSTR* Error: Any messages if errors occur
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Add a new pest for this dose calculation.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the add.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(SetDosePest)( /*[in]*/ int PestIndex, 
                           /*[in]*/ int FumigationType, 
                           /*[in]*/ int LifeStage, 
                           /*[in]*/ int TreatmentType,
                           /*[in]*/ int Commodity,
                           /*[in]*/ double HLT, 
                           /*[in]*/ double Temp, 
                           /*[in]*/ double TimeExposure, 
                           /*[in]*/ double Volume,
						   /*[in]*/ double LoadFactor,
						   /*[in]*/ int IsComplete,
						   /*[in]*/ int ICounter,
                           /*[out]*/ BSTR* Error );

   // =DOC SECTION Method GetPestIndex( /*[in]*/ BSTR* Pest, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC SECTION Get the pest index based on the pest name.
   // =DOC TEXT public  GetPestIndex( /*[in]*/ BSTR* Pest, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC TEXT /*[in]*/ BSTR* Pest: The pest name
   // =DOC TEXT /*[out, retval]*/ int* Index: The pest index
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the pest index based on the pest name.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the pest.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetPestIndex)( /*[in]*/ BSTR* Pest, 
                            /*[out, retval]*/ int* Index );

   // =DOC SECTION Method GetPestString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* Pest ): STDMETHOD
   // =DOC SECTION Get the pest name based on the pest index.
   // =DOC TEXT public  GetPestString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* Pest ): STDMETHOD
   // =DOC TEXT /*[in]*/ int Index: The pest index
   // =DOC TEXT /*[out, retval]*/ BSTR* Pest: The pest name
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the pest name based on the pest index.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the pest.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetPestString)( /*[in]*/ int Index, 
                             /*[out, retval]*/ BSTR* Pest );

   // =DOC SECTION Method GetPestTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC SECTION Get the total number of pests supported by the calculator.
   // =DOC TEXT public  GetPestTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ int* Total: The total number of pests
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the total number of pests supported by the calculator.  
   // =DOC TEXT Return S_OK.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetPestTotal)( /*[out, retval]*/ int* Total );

   // =DOC SECTION Method GetCommodityIndex( /*[in]*/ BSTR* Commodity, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC SECTION Get the commodity index based on the commodity name.
   // =DOC TEXT public  GetCommodityIndex( /*[in]*/ BSTR* Commodity, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC TEXT /*[in]*/ BSTR* Commodity: The commodity name
   // =DOC TEXT /*[out, retval]*/ int* Index: The commodity index
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the commodity index based on the commodity name.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the commodity.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetCommodityIndex)( /*[in]*/ BSTR *Commodity, 
                                 /*[out, retval]*/ int *Index );

   // =DOC SECTION Method GetCommodityString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* Commodity ): STDMETHOD
   // =DOC SECTION Get the commodity name based on the commodity index.
   // =DOC TEXT public  GetCommodityString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* Commodity ): STDMETHOD
   // =DOC TEXT /*[in]*/ int Index: The commodity index
   // =DOC TEXT /*[out, retval]*/ BSTR* Commodity: The commodity name
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the commodity name based on the commodity index.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the commodity.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetCommodityString)( /*[in]*/ int Index, 
                                  /*[out, retval]*/ BSTR *Commodity );

   // =DOC SECTION Method GetCommodityTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC SECTION Get the total number of commodities supported by the calculator.
   // =DOC TEXT public  GetCommodityTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ int* Total: The total number of commodities
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the total number of commodities supported by the calculator.  
   // =DOC TEXT Return S_OK.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(GetCommodityTotal)( /*[out, retval]*/ int *Total );
   
   // =DOC SECTION Method GetTreatmentTypeString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* TreatmentType ): STDMETHOD
   // =DOC SECTION Get the treatment type name based on the treatment type index.
   // =DOC TEXT public  GetTreatmentTypeString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* TreatmentType ): STDMETHOD
   // =DOC TEXT /*[in]*/ int Index: The treatment type index
   // =DOC TEXT /*[out, retval]*/ BSTR* TreatmentType: The treatment type name
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the treatment type name based on the treatment type index.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the treatment type.
   //////////////////////////////////////////////////////////////////
   STDMETHOD(GetTreatmentTypeString)( /*[in]*/ int Index, 
                                      /*[out, retval]*/ BSTR* TreatmentType );
   
   // =DOC SECTION Method GetTreatmentTypeIndex( /*[in]*/ BSTR* TreatmentType, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC SECTION Get the treatment type index based on the treatment type name.
   // =DOC TEXT public  GetTreatmentTypeIndex( /*[in]*/ BSTR* TreatmentType, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC TEXT /*[in]*/ BSTR* TreatmentType: The treatment type name
   // =DOC TEXT /*[out, retval]*/ int* Index: The treatment type index
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the treatment type index based on the treatment type name.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the treatment type.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetTreatmentTypeIndex)( /*[in]*/ BSTR* TreatmentType, 
                                     /*[out, retval]*/ int* Index );
   
   // =DOC SECTION Method GetTreatmentTypeTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC SECTION Get the total number of treatment types supported by the calculator.
   // =DOC TEXT public  GetTreatmentTypeTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ int* Total: The total number of treatment types
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the total number of treatment types supported by the calculator.  
   // =DOC TEXT Return S_OK.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetTreatmentTypeTotal)( /*[out, retval]*/ int* Total );
   
   // =DOC SECTION Method GetFumigationTypeString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* FumigationType ): STDMETHOD
   // =DOC SECTION Get the fumigation type name based on the fumigation type index.
   // =DOC TEXT public  GetFumigationTypeString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* FumigationType ): STDMETHOD
   // =DOC TEXT /*[in]*/ int Index: The fumigation type index
   // =DOC TEXT /*[out, retval]*/ BSTR* FumigationType: The fumigation type name
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the fumigation type name based on the fumigation type index.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the fumigation type.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetFumigationTypeString)( /*[in]*/ int Index, 
                                       /*[out, retval]*/ BSTR* FumigationType );
   
   // =DOC SECTION Method GetFumigationTypeIndex( /*[in]*/ BSTR* FumigationType, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC SECTION Get the fumigation type index based on the fumigation type name.
   // =DOC TEXT public  GetFumigationTypeIndex( /*[in]*/ BSTR* FumigationType, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC TEXT /*[in]*/ BSTR* FumigationType: The fumigation type name
   // =DOC TEXT /*[out, retval]*/ int* Index: The fumigation type index
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the fumigation type index based on the fumigation type name.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the fumigation type.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetFumigationTypeIndex)( /*[in]*/ BSTR* FumigationType, 
                                      /*[out, retval]*/ int* Index );
   
   // =DOC SECTION Method GetFumigationTypeTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC SECTION Get the total number of fumigation types supported by the calculator.
   // =DOC TEXT public  GetFumigationTypeTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ int* Total: The total number of fumigation types
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the total number of fumigation types supported by the calculator.  
   // =DOC TEXT Return S_OK.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetFumigationTypeTotal)( /*[out, retval]*/ int* Total );
   
   // =DOC SECTION Method GetLifeStageString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* LifeStage ): STDMETHOD
   // =DOC SECTION Get the life stage name based on the life stage index.
   // =DOC TEXT public  GetLifeStageString( /*[in]*/ int Index, /*[out, retval]*/ BSTR* LifeStage ): STDMETHOD
   // =DOC TEXT /*[in]*/ int Index: The life stage index
   // =DOC TEXT /*[out, retval]*/ BSTR* LifeStage: The life stage name
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the life stage name based on the life stage index.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the life stage.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetLifeStageString)( /*[in]*/ int Index, 
                                  /*[out, retval]*/ BSTR* LifeStage );
   
   // =DOC SECTION Method GetLifeStageIndex( /*[in]*/ BSTR* LifeStage, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC SECTION Get the life stage index based on the life stage name.
   // =DOC TEXT public  GetLifeStageIndex( /*[in]*/ BSTR* LifeStage, /*[out, retval]*/ int* Index ): STDMETHOD
   // =DOC TEXT /*[in]*/ BSTR* LifeStage: The life stage name
   // =DOC TEXT /*[out, retval]*/ int* Index: The life stage index
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the life stage index based on the life stage name.  
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems finding the life stage.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetLifeStageIndex)( /*[in]*/ BSTR* LifeStage, 
                                 /*[out, retval]*/ int* Index );

   // =DOC SECTION Method GetLifeStageTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC SECTION Get the total number of life stages supported by the calculator.
   // =DOC TEXT public  GetLifeStageTotal( /*[out, retval]*/ int* Total ): STDMETHOD
   // =DOC TEXT /*[out, retval]*/ int* Total: The total number of life stages
   // =DOC TEXT Returns HRESULT: Standard HRESULT values
   // =DOC TEXT Get the total number of life stages supported by the calculator.  
   // =DOC TEXT Return S_OK.
   //////////////////////////////////////////////////////////////////
	STDMETHOD(GetLifeStageTotal)( /*[out, retval]*/ int* Total );



protected:
	void CleanTreatmentType();
	void CleanLifeStage();
	void CleanFumigationType();
	// =DOC SECTION Method RemoveDuplicateWarning(UINT unWarningId, CString& strNewWarning, CString& strExistingWarning): void
	// =DOC SECTION Removes duplicate warnings that may already be in the existing warning string from the new warning string.
	// =DOC TEXT protected  RemoveDuplicateWarning(UINT unWarningId, CString& strNewWarning, CString& strExistingWarning): void
	// =DOC TEXT Returns void:
   // =DOC TEXT UINT unWarningId: The warning string ID to search for
   // =DOC TEXT CString& strNewWarning: The new warning string
   // =DOC TEXT CString& strExistingWarning: The existing warning string

	// =DOC TEXT Removes duplicate warnings that may already be in the existing warning string from the new warning string.
	//////////////////////////////////////////////////////////////////
	void RemoveDuplicateWarning( UINT unWarningId, 
                                CString& strNewWarning, 
                                CString& strExistingWarning );

	// =DOC SECTION Method CleanPest(): void
	// =DOC SECTION Clears the pest names from m_astrPest.
	// =DOC TEXT protected  CleanPest(): void
	// =DOC TEXT Returns void:
	// =DOC TEXT Clears the pest strings from m_astrPest.
	//////////////////////////////////////////////////////////////////
	void CleanPest();

	// =DOC SECTION Method PopulatePest(): void
	// =DOC SECTION Populates the pest names for m_astrPest.
	// =DOC TEXT protected  PopulatePest(): void
	// =DOC TEXT Returns void: 
	// =DOC TEXT Populates the pest strings for m_astrPest.
	//////////////////////////////////////////////////////////////////
	void PopulatePest();

	// =DOC SECTION Method CleanCommodity(): void
	// =DOC SECTION Clears the commodity names from m_astrCommodity.
	// =DOC TEXT protected  CleanCommodity(): void
	// =DOC TEXT Returns void:
	// =DOC TEXT Clears the commodity names from m_astrCommodity.
	//////////////////////////////////////////////////////////////////
	void CleanCommodity();

	// =DOC SECTION Method PopulateCommodity(): void
	// =DOC SECTION Populates the commodity names for m_astrCommodity.
	// =DOC TEXT protected  PopulateCommodity(): void
	// =DOC TEXT Returns void: 
	// =DOC TEXT Populates the commodity strings for m_astrCommodity.
	//////////////////////////////////////////////////////////////////
	void PopulateCommodity();

	// =DOC SECTION Method PopulateLifeStage(): void
	// =DOC SECTION Populates the life stage names for m_astrLifeStage.
	// =DOC TEXT protected  PopulateLifeStage(): void
	// =DOC TEXT Returns void: 
	// =DOC TEXT Populates the life stage strings for m_astrLifeStage.
	//////////////////////////////////////////////////////////////////
	void PopulateLifeStage();
	
   // =DOC SECTION Method PopulateTreatmentType(): void
	// =DOC SECTION Populates the treatment type names for m_astrTreatmentType.
	// =DOC TEXT protected  PopulateTreatmentType(): void
	// =DOC TEXT Returns void: 
	// =DOC TEXT Populates the treatment type strings for m_astrTreatmentType.
	//////////////////////////////////////////////////////////////////
	void PopulateTreatmentType();
	
   // =DOC SECTION Method PopulateFumigationType(): void
	// =DOC SECTION Populates the fumigation type names for m_astrFumigationType.
	// =DOC TEXT protected  PopulateFumigationType(): void
	// =DOC TEXT Returns void: 
	// =DOC TEXT Populates the fumigation type strings for m_astrFumigationType.
	//////////////////////////////////////////////////////////////////
	void PopulateFumigationType();

   // Member variables

	// =DOC SECTION Member m_pcalcMonPt: DMonitorPoint*
	// =DOC SECTION Monitor point calculator
	// =DOC TEXT m_pcalcMonPt: DMonitorPoint*
	// =DOC TEXT Monitor point calculator
	//////////////////////////////////////////////////////////////////
	DMonitorPoint* m_pcalcMonPt;

	// =DOC SECTION Member m_pvalidateHalfLifeInput: DValidateHalfLifeInput*
	// =DOC SECTION Half loss validator
	// =DOC TEXT m_pvalidateHalfLifeInput: DValidateHalfLifeInput*
	// =DOC TEXT Half loss validator
	//////////////////////////////////////////////////////////////////
	DValidateHalfLifeInput* m_pvalidateHalfLifeInput;

	// =DOC SECTION Member m_nMaxProfumeIndex: int
	// =DOC SECTION The index of the pest requiring the maximum amount of ProFume.
	// =DOC TEXT m_nMaxProfumeIndex: int
	// =DOC TEXT The index of the pest requiring the maximum amount of ProFume.
	//////////////////////////////////////////////////////////////////
	int m_nMaxProfumeIndex;

	// =DOC SECTION Member m_astrPest: CStringArray
	// =DOC SECTION An array of pest names.
	// =DOC TEXT m_astrPest: CStringArray
	// =DOC TEXT An array of pest names.
	//////////////////////////////////////////////////////////////////
	CStringArray m_astrPest;

	// =DOC SECTION Member m_astrCommodity: CStringArray
	// =DOC SECTION An array of commodity names.
	// =DOC TEXT m_astrCommodity: CStringArray
	// =DOC TEXT An array of commodity names.
	//////////////////////////////////////////////////////////////////
	CStringArray m_astrCommodity;

	// =DOC SECTION Member m_astrLifeStage: CStringArray
	// =DOC SECTION An array of life stage names.
	// =DOC TEXT m_astrLifeStage: CStringArray
	// =DOC TEXT An array of life stage names.
	//////////////////////////////////////////////////////////////////
	CStringArray m_astrLifeStage;

	// =DOC SECTION Member m_astrFumigationType: CStringArray
	// =DOC SECTION An array of fumigation type names.
	// =DOC TEXT m_astrFumigationType: CStringArray
	// =DOC TEXT An array of fumigation type names.
	//////////////////////////////////////////////////////////////////
	CStringArray m_astrFumigationType;

	// =DOC SECTION Member m_astrTreatmentType: CStringArray
	// =DOC SECTION An array of treatment type names.
	// =DOC TEXT TreatmentType: CStringArray
	// =DOC TEXT An array of treatment type names.
	//////////////////////////////////////////////////////////////////
	CStringArray m_astrTreatmentType;

   // =DOC SECTION Member m_padataDose: DOSE_DATA_ARRAY*
   // =DOC SECTION An array of dose information for each pest.
   // =DOC TEXT m_padataDose: DOSE_DATA_ARRAY*
   // =DOC TEXT An array of dose information for each pest.
   //////////////////////////////////////////////////////////////////
   DOSE_DATA_ARRAY* m_padataDose;

   // =DOC SECTION Member m_pvalidateDoseInput: DValidateDoseInput*
   // =DOC SECTION Dose user input validator.
   // =DOC TEXT m_pvalidateDoseInput: DValidateDoseInput*
   // =DOC TEXT Dose user input validator.
   //////////////////////////////////////////////////////////////////
   DValidateDoseInput* m_pvalidateDoseInput;

   // =DOC SECTION Member m_pcalcDose: DDoseCalc*
   // =DOC SECTION Dose calculator.
   // =DOC TEXT m_pcalcDose: DDoseCalc*
   // =DOC TEXT Dose calculator.
   //////////////////////////////////////////////////////////////////
   DDoseCalc* m_pcalcDose;

   // =DOC SECTION Member m_padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC SECTION An array of all the concentration data.
   // =DOC TEXT m_padataConc: CONCENTRATION_DATA_ARRAY*
   // =DOC TEXT An array of all the concentration data.
   //////////////////////////////////////////////////////////////////
   CONCENTRATION_DATA_ARRAY* m_padataConc;

   // =DOC SECTION Member m_pvalidateConc: DValidateConcentration*
   // =DOC SECTION Concentration validator.
   // =DOC TEXT m_pvalidateConc: DValidateConcentration*
   // =DOC TEXT Concentration validator.
   //////////////////////////////////////////////////////////////////
   DValidateConcentration* m_pvalidateConc;

   // =DOC SECTION Member m_padataAreaMonPt: AREA_MON_PT_DATA_ARRAY*
   // =DOC SECTION An array of all the area monitor point data.
   // =DOC TEXT m_padataAreaMonPt: AREA_MON_PT_DATA_ARRAY*
   // =DOC TEXT An array of all the area monitor point data.
   //////////////////////////////////////////////////////////////////
   AREA_MON_PT_DATA_ARRAY* m_padataAreaMonPt;

   // =DOC SECTION Member m_pvalidateAreaMonPt: DValidateAreaMonPtInput*
   // =DOC SECTION Area monitor point validator.
   // =DOC TEXT m_pvalidateAreaMonPt: DValidateAreaMonPtInput*
   // =DOC TEXT Area monitor point validator.
   //////////////////////////////////////////////////////////////////
   DValidateAreaMonPtInput* m_pvalidateAreaMonPt;

	// =DOC SECTION Member m_pcalcArea: DArea*
	// =DOC SECTION Area calculator
	// =DOC TEXT m_pcalcArea: DArea*
	// =DOC TEXT Area calculator
	//////////////////////////////////////////////////////////////////
   DArea* m_pcalcArea;

	// Jett 6/10/03 - added the C0 variables so we can use it in 
    // Part II when we call GetMonitorPointStatus

   	// =DOC SECTION Member m_dInitConc: double
	// =DOC SECTION interpolated C0 value
	// =DOC TEXT m_dInitConc: double
	// =DOC TEXT interpolated C0 value
	//////////////////////////////////////////////////////////////////
   double m_dInitConc;

};

#endif //__PROFUMECALCULATOR_H_
