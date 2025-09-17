// =DOC FILE Global constants for dpcalc
// =DOC TEXT Declares global constants for dpcalc 
//////////////////////////////////////////////////////////////////

// GlobalEnums.h: declarations of Global constants/enumerations
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_GLOBALENUMS_H__CF7E9B4C_F8A0_4559_93C8_62282677BEF4__INCLUDED_)
#define AFX_GLOBALENUMS_H__CF7E9B4C_F8A0_4559_93C8_62282677BEF4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//JNM 04/01/03 - Fumiguide Enhancements 2003 - Modified the Inside Diameter constants.
const double SHOOTINGLINE_DIAMETER_ONE   = 3.175;
const double SHOOTINGLINE_DIAMETER_TWO   = 4.37;
const double SHOOTINGLINE_DIAMETER_THREE = 6.35;

const double MIN_FAN = 28.31;
const double MAX_FAN = 707.92;

const double MIN_TEMP = 4.44;
const double MAX_TEMP = 65.56;
//Pat 4/7/2005 - Added these values for pressure
const double MIN_PRES = 0.0;
const double MAX_PRES = 25.0;

const double MIN_LINE_LENGTH = 0.0;
const double MAX_LINE_LENGTH = 152.0;

//Pat 1/06/2005 - Added new constants for new fields
const double MIN_ONEFOURTH_LENGTH = 0.0;
const double MAX_ONEFOURTH_LENGTH = 213.36;
const double MIN_ONEEIGHT_LENGTH = 0.0;
const double MAX_ONEEIGHT_LENGTH = 30.48;

const double MIN_SHOOTING_RATE = 0.0;
const double MAX_SHOOTING_RATE = 9.0;

const double MIN_PEST_TEMP = 20.0; // Minimum temperature where a warning is required

const double MIN_HLT  = 0.0;

const double MIN_LOAD_FACTOR = 0.0; //PAT 1/17/06 - Added

const double MIN_TIME = 0.0;

const double MIN_VOL  = 0.0;

const double MAX_CONCENTRATION = 128;        // The maximum concentration = 128 mg/liter
const double MAX_CONCENTRATION_TIME = 1500;  // The maximum concentration time = 1500 mg-hr/liter
const double KILO_PER_CYL = 56.7;            // 125 lbs/cylinder * 0.4536 kg/lb = 56.7 kg/cylinder

// Safety factors for different life stages
const double DOSE_SAFETY_FACTOR_EGGS_SPACE = 1.0;
const double DOSE_SAFETY_FACTOR_EGGS_COMMODITY = 1.5;
const double DOSE_SAFETY_FACTOR_POST_EMBRYONIC = 1.5;
const double DOSE_SAFETY_FACTOR_POST_EMBRYONIC_INSIDE_RANGE = 2.0;
const double DOSE_SAFETY_FACTOR_POST_EMBRYONIC_OUTSIDE_RANGE = 2.5;

const double RANGE_TIME_START = 18.0;
const double RANGE_TIME_END = 32.0;

// Probit values from DAS
const double PROBIT_99 = 2.32635;
const double PROBIT_95 = 1.64485;
const double PROBIT_90 = 1.28155;
const double PROBIT_85 = 1.03643;
const double PROBIT_80 = 0.84162;
const double PROBIT_75 = 0.67449;
const double PROBIT_70 = 0.5244;
const double PROBIT_65 = 0.38532;
const double PROBIT_60 = 0.25335;
const double PROBIT_55 = 0.12566;
const double PROBIT_50 = 0.0;

const double TARGET_RANGE = 0.02;   // 2% of target

// Valid elapsed time difference
const double ELAPSED_TIME_DIFF = 0.5;

//CG 08/28/2003 Added constant for Delta T (sub A) which will be used in Step II 9 B, C and E
// This represents the amt of time that it takes to reach a new equilibrium concentration
// once additional gas has been injected
const double DELTA_TIME_A = 0.5;


const short  MIN_HUMIDITY = 0;
const short  MAX_HUMIDITY = 100;

const int PROFUME_REINTRO_TIME_PADDING = 1;  // ProFume reintroduction time padding

// Status value
const int STATUS_UNKNOWN = 0;
const int STATUS_ON_TARGET = 1;
const int STATUS_ABOVE_TARGET = 2;
const int STATUS_BELOW_TARGET = 3;
const int STATUS_ACHIEVED_TARGET = 4;
const int STATUS_INCREASING_CONCENTRATION = 5;
const int STATUS_TARGET_DOSE_INCOMPATABLE = 6;
const int STATUS_BAD_ELAPSED_TIMES = 7;
const int STATUS_MISSING_CONC_READING = 8;


#endif // !defined(AFX_GLOBALENUMS_H__CF7E9B4C_F8A0_4559_93C8_62282677BEF4__INCLUDED_)
