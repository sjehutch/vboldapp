// =DOC FILE Generic global functions for dpcalc
// =DOC TEXT Declares calculator functions that are used in multiple classes
//////////////////////////////////////////////////////////////////

// GlobalFunctions.h: interface the generic global functions in the dpcalc project
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_GLOBALFUNCTIONS_H__4358C067_7F95_47A0_89A6_A9768166C43E__INCLUDED_)
#define AFX_GLOBALFUNCTIONS_H__4358C067_7F95_47A0_89A6_A9768166C43E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



// =DOC SECTION Method GetFinalDose( const double &dIntercept, const double &dSlope, const double &dProbit, const double &dSafetyFactor ): double
// =DOC SECTION Get the final dose of ProFume required for a fumigation
// =DOC TEXT GetFinalDose( const double &dIntercept, const double &dSlope, const double &dProbit, const double &dSafetyFactor ): double
// =DOC TEXT double &dIntercept: Pest's intercept value
// =DOC TEXT double &dSlope: Pest's slope value
// =DOC TEXT double &dProbit: Pest's probit value
// =DOC TEXT double &dSafetyFactor: Pest's safety factor value
// =DOC TEXT Returns double: Final dose value
// =DOC TEXT Get the final dose of ProFume required for a fumigation.  
//////////////////////////////////////////////////////////////////
double GetFinalDose( const double &dIntercept, const double &dSlope, const double &dProbit, const double &dSafetyFactor );


// =DOC SECTION Method GetInitialConcentration( const double &dHltEstimated, const double &dTimeExposure, const double &dDoseFinal, const double &dNValue ): double
// =DOC SECTION Get the initial concentration of ProFume based on current data (g/m^3)
// =DOC TEXT GetFutureConcentration( const double &dHltEstimated, const double &dTimeExposure, const double &dDoseFinal, const double &dNValue ): double
// =DOC TEXT double &dHltEstimated: Estimated HLT (hrs)
// =DOC TEXT double &dTimeExposure: Time of exposure (hrs)
// =DOC TEXT double &dDoseFinal: Final dose (g/m^3)
// =DOC TEXT double &dNValue: Pest's N value
// =DOC TEXT Returns double: Initial concentration value (g/m^3)
// =DOC TEXT Get the initial concentration of ProFume based on current data (g/m^3)
//////////////////////////////////////////////////////////////////
double GetInitialConcentration( const double &dHltEstimated, 
                                const double &dTimeExposure, 
                                const double &dDoseFinal, 
                                const double &dNValue );

//12/22-04-YYT-Added Get initial concentration 2 when user has defined a target CT
// =DOC SECTION Method GetInitialConcentration2( const double &dHltOriginal, const double &dTimeExposure, const double &UserTargetCT, const double &dNValue): double
// =DOC SECTION Get the initial concentration of ProFume based on current data (g/m^3)
// =DOC TEXT GetFutureConcentration( const double &dHltOriginal, const double &dTimeExposure, const double &UserTargetCT, const double &dNValue): double
// =DOC TEXT double &dHltOriginal: Original HLT (hrs)
// =DOC TEXT double &dTimeExposure: Time of exposure (hrs)
// =DOC TEXT double &dNValue: Pest's N value
// =DOC TEXT double &UserTargetCT: User-defined Target CT
// =DOC TEXT Returns double: Initial concentration value (g/m^3)
// =DOC TEXT Get the initial concentration of ProFume based on current data (g/m^3)
//////////////////////////////////////////////////////////////////
double GetInitialConcentration2( const double &dHltOriginal, 
                                const double &dTimeExposure, 
                                const double &UserTargetCT, 
                                const double &dNValue );

// =DOC SECTION Method GetEHLT(double dOHLT, double dNValue): double
// =DOC SECTION Convert OHLT to EHLT.
// =DOC TEXT GetEHLT(double dOHLT, double dNValue): double
// =DOC TEXT dOHLT: The original half-loss time to convert.
// =DOC TEXT dNValue: The N-value to use in the conversion.
// =DOC TEXT Returns double: The exponent-HLT.
//////////////////////////////////////////////////////////////////
double GetEHLT(double dOHLT, double dNValue);


#endif // !defined(AFX_GLOBALFUNCTIONS_H__4358C067_7F95_47A0_89A6_A9768166C43E__INCLUDED_)
