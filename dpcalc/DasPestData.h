// DasPestData.h: interface for the DDasPestData class.
// JNM 04/09/03 - Fumiguide Enh. 2003 - Added 2 new pests.
//									  - Added function declarations for the 
//										2 new pests.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DASPESTDATA_H__3CC3C604_6FE2_438F_8592_BD1A0E63A310__INCLUDED_)
#define AFX_DASPESTDATA_H__3CC3C604_6FE2_438F_8592_BD1A0E63A310__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "PestData.h"
#include "CommodityData.h"
#include "DasFumigationTypeData.h"
#include "DasLifeStageData.h"
#include "DasTreatmentTypeData.h"


// =DOC SECTION enum PEST
// =DOC SECTION Pest enumeration
// =DOC TEXT enum PEST extends nothing
// =DOC TEXT Pest enumeration
// =DOC TEXT PEST_CONFUSED_FLOUR_BEETLE    = 0
// =DOC TEXT PEST_RED_FLOUR_BEETLE         = 1
// =DOC TEXT PEST_SAWTOOTHED_GRAIN_BEETLE  = 2
// =DOC TEXT PEST_WAREHOUSE_BEETLE         = 3
// =DOC TEXT PEST_INDIAN_MEAL_MOTH         = 4
// =DOC TEXT PEST_MEDITERRANEAN_FLOUR_MOTH = 5
// =DOC TEXT PEST_CODLING_MOTH             = 6
// =DOC TEXT PEST_NAVEL_ORANGEWORM         = 7
// =DOC TEXT PEST_OTHER_BEETLE             = 8
// =DOC TEXT PEST_OTHER_MOTH               = 9
// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added these 2 new pests.
// =DOC TEXT PEST_GRANARY_WEEVIL		   = 10
// =DOC TEXT PEST_RICE_WEEVIL			   = 11
// Pat 4/6/05: Added new Pest
// =DOC TEXT PEST_LESSER_GRAIN_BORER       = 12
// Pat 12/1/2005 - Enhancement 2005-2006 - Added new pests 
// =DOC TEXT PEST_HIDE_BEETLE			  = 13
// =DOC TEXT PEST_DRUG_STORE_BEETLE		  = 14
// =DOC TEXT PEST_RUSTY_GRAIN_BEETLE	  = 15
// =DOC TEXT PEST_ALMOND_MOTH		   	  = 16
// =DOC TEXT PEST_RODENTS			      = 17
//////////////////////////////////////////////////////////////////
enum PEST { PEST_CONFUSED_FLOUR_BEETLE    = 0,
            PEST_RED_FLOUR_BEETLE         = 1,
            PEST_SAWTOOTHED_GRAIN_BEETLE  = 2,
            PEST_WAREHOUSE_BEETLE         = 3,
            PEST_INDIAN_MEAL_MOTH         = 4,
            PEST_MEDITERRANEAN_FLOUR_MOTH = 5,
            PEST_CODLING_MOTH             = 6,
            PEST_NAVEL_ORANGEWORM         = 7,
            PEST_OTHER_BEETLE             = 8,
            PEST_OTHER_MOTH               = 9,
			PEST_GRANARY_WEEVIL			  = 10,
			PEST_RICE_WEEVIL			  = 11,
			PEST_LESSER_GRAIN_BORER       = 12,
			PEST_HIDE_BEETLE			  = 13, 
			PEST_DRUG_STORE_BEETLE		  = 14, 
			PEST_RUSTY_GRAIN_BEETLE		  = 15, 
			PEST_ALMOND_MOTH		   	  = 16, 
			PEST_RODENTS			      = 17,
			PEST_CIGARETTE_BEETLE		  = 18,
			PEST_COWPEA_WEEVIL            = 19,
			PEST_BEAN_WEEVIL              = 20,
			PEST_COCOA_MOTH 			  = 21};


// =DOC SECTION Class DDasPestData
// =DOC SECTION DAS pest data handler
// =DOC TEXT class DDasPestData extends nothing.
// =DOC TEXT The class contains the pest data used in all algorithms.
//////////////////////////////////////////////////////////////////
class DDasPestData  
{
public:
   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DDasPestData();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
   virtual ~DDasPestData();


   // =DOC SECTION Method GetPestNames( CStringArray &astrPestNames ): void
	// =DOC SECTION Gets the pest names and places them in the astrPestNames parameter 
   // =DOC TEXT public  GetPestNames( CStringArray &astrPestNames ): void
   // =DOC TEXT CStringArray &astrPestNames: Array of pest names
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves a list of pest names in which data exists
   //////////////////////////////////////////////////////////////////
	void GetPestNames( CStringArray &astrPestNames );


   // =DOC SECTION Method PopulatePestData( PEST &ePest, LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populates the internal member m_aPestData for a particular pest 
   // =DOC TEXT public  PopulatePestData( PEST &ePest, LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
   // =DOC TEXT PEST &ePest: The pest
   // =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
   // =DOC TEXT The functions retrieves the lower/upper temp bound pest data for a requested temperature
   //////////////////////////////////////////////////////////////////
   void PopulatePestData( const PEST &ePest, 
                          const LIFE_STAGE &eLifeStage, 
                          const TREATMENT_TYPE &eTreatmentType,
                          const FUMIGATION_TYPE &eFumigationType,
                          const double &dTimeExposure,
                          CString &strMsg );


   // =DOC SECTION Method GetPestData( double &dTemp, BOOL bUpperValue ): PEST_DATA
// =DOC SECTION Retrieves the lower/upper temp bound pest data for a requested temperature
   // =DOC TEXT public  GetPestData( double &dTemp, BOOL bUpperValue ): PEST_DATA
   // =DOC TEXT double dTemp: Temperature for the fumigation
   // =DOC TEXT BOOL bUpperValue: If dTemp is between temperatures, return the higher temps pest data
   // =DOC TEXT Returns PEST_DATA: Structure containing pest data for a particular temperature
   // =DOC TEXT The functions retrieves the lower/upper temp bound pest data for a requested temperature
   //////////////////////////////////////////////////////////////////
   PEST_DATA GetPestData( const double &dTemp, BOOL bUpperValue );
	
protected:

	// =DOC SECTION Method PopulateRedFlourBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with red flour beetle data
	// =DOC TEXT protected  PopulateRedFlourBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with red flour beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateRedFlourBeetle( const LIFE_STAGE &eLifeStage, 
                                const TREATMENT_TYPE &eTreatmentType, 
                                const FUMIGATION_TYPE &eFumigationType,
                                const double &dTimeExposure,
                                CString &strMsg );

	// =DOC SECTION Method PopulateSawToothedGrainBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with saw toothed grain beetle data
	// =DOC TEXT protected  PopulateSawToothedGrainBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with saw toothed grain beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateSawToothedGrainBeetle( const LIFE_STAGE &eLifeStage, 
                                       const TREATMENT_TYPE &eTreatmentType, 
                                       const FUMIGATION_TYPE &eFumigationType,
                                       const double &dTimeExposure,
                                       CString &strMsg );

	// =DOC SECTION Method PopulateMedFlourMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with mediteranean flour moth data
	// =DOC TEXT protected  PopulateMedFlourMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with mediteranean flour moth data
	//////////////////////////////////////////////////////////////////
	void PopulateMedFlourMoth( const LIFE_STAGE &eLifeStage, 
                              const TREATMENT_TYPE &eTreatmentType, 
                              const FUMIGATION_TYPE &eFumigationType,
                              const double &dTimeExposure,
                              CString &strMsg );

	// =DOC SECTION Method PopulateIndianMealMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with indian meal moth data
	// =DOC TEXT protected  PopulateIndianMealMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with indian meal moth data
	//////////////////////////////////////////////////////////////////
	void PopulateIndianMealMoth( const LIFE_STAGE &eLifeStage, 
                                const TREATMENT_TYPE &eTreatmentType, 
                                const FUMIGATION_TYPE &eFumigationType,
                                const double &dTimeExposure,
                                CString &strMsg );

	// =DOC SECTION Method PopulateConfusedFlourBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with confused flour beetle data
	// =DOC TEXT protected  PopulateConfusedFlourBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with confused flour beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateConfusedFlourBeetle( const LIFE_STAGE &eLifeStage, 
                                     const TREATMENT_TYPE &eTreatmentType, 
                                     const FUMIGATION_TYPE &eFumigationType,
                                     const double &dTimeExposure,
                                     CString &strMsg );

	// =DOC SECTION Method PopulateWarehouseBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateWarehouseBeetle( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new method for new pest.
	// =DOC SECTION Method PopulateGranaryWeevil( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with granary weevil data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with granary weevil data
	//////////////////////////////////////////////////////////////////
	void PopulateGranaryWeevil( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// JNM 04/01/03 - Fumiguide Enhancements 2003 - Added new method for new pest.
	// =DOC SECTION Method PopulateGranaryWeevil( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with rice weevil data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with rice weevil data
	//////////////////////////////////////////////////////////////////
	void PopulateRiceWeevil( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// =DOC SECTION Method PopulateCodlingMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with codling moth data
	// =DOC TEXT protected  PopulateCodlingMoth( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with codling moth data
	void PopulateCodlingMoth( const LIFE_STAGE &eLifeStage, 
                             const TREATMENT_TYPE &eTreatmentType, 
                             const FUMIGATION_TYPE &eFumigationType,
                             const double &dTimeExposure,
                             CString &strMsg );

	// =DOC SECTION Method PopulateNavelOrangeWorm( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with navel orange worm data
	// =DOC TEXT protected  PopulateNavelOrangeWorm( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
   // =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
   // =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
   // =DOC TEXT double &dTimeExposure: The time of exposure
   // =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with navel orange worm data
	//////////////////////////////////////////////////////////////////
	void PopulateNavelOrangeWorm( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	
	// Pat 4/6/05 : Added function for new pest
	// =DOC SECTION Method PopulateLesserGrainBorer( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateLesserGrainBorer( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateHideBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateHideBeetle( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateDrugStoreBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateDrugStoreBeetle( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRustyGrainBeetle( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateRustyGrainBeetle( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateAlmondMoth ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateAlmondMoth ( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRodents ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateRodents ( const LIFE_STAGE &eLifeStage, 
                                 const TREATMENT_TYPE &eTreatmentType, 
                                 const FUMIGATION_TYPE &eFumigationType,
                                 const double &dTimeExposure,
                                 CString &strMsg );



	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRodents ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateCigarette_Bettle ( const LIFE_STAGE &eLifeStage, 
									 const TREATMENT_TYPE &eTreatmentType, 
									 const FUMIGATION_TYPE &eFumigationType,
									 const double &dTimeExposure,
									 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRodents ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateCowpeaWeevil ( const LIFE_STAGE &eLifeStage, 
									 const TREATMENT_TYPE &eTreatmentType, 
									 const FUMIGATION_TYPE &eFumigationType,
									 const double &dTimeExposure,
									 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRodents ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateBeanWeevil ( const LIFE_STAGE &eLifeStage, 
									 const TREATMENT_TYPE &eTreatmentType, 
									 const FUMIGATION_TYPE &eFumigationType,
									 const double &dTimeExposure,
									 CString &strMsg );

	// Pat 12/1/05 : Enhancement 2005-2006 Added function for new pests
	// =DOC SECTION Method PopulateRodents ( LIFE_STAGE &eLifeStage, TREATMENT_TYPE &eTreatmentType, CString &strMsg ): void
	// =DOC SECTION Populate m_aPestData with warehouse beetle data
	// =DOC TEXT LIFE_STAGE &eLifeStage: The life stage of the pest
	// =DOC TEXT TREATMENT_TYPE &eTreatmentType: The treatment type of the pest
	// =DOC TEXT FUMIGATION_TYPE &eFumigationType: The fumigation type
	// =DOC TEXT double &dTimeExposure: The time of exposure
	// =DOC TEXT CString &strMsg: Error message if problems occur
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Populate a pest data array (m_aPestData) with warehouse beetle data
	//////////////////////////////////////////////////////////////////
	void PopulateCocoaMoth ( const LIFE_STAGE &eLifeStage, 
									 const TREATMENT_TYPE &eTreatmentType, 
									 const FUMIGATION_TYPE &eFumigationType,
									 const double &dTimeExposure,
									 CString &strMsg );












	
   
   // =DOC SECTION Member m_aPestData: PEST_DATA_ARRAY
   // =DOC SECTION Array of pest data
   // =DOC TEXT m_aPestData: PEST_DATA_ARRAY
   // =DOC TEXT Array of pest data populated by the Populate* functions
   //////////////////////////////////////////////////////////////////
   PEST_DATA_ARRAY m_aPestData;
};

#endif // !defined(AFX_DASPESTDATA_H__3CC3C604_6FE2_438F_8592_BD1A0E63A310__INCLUDED_)
