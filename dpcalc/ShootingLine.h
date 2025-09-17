// =DOC FILE Shooting Line interface class
// =DOC TEXT Shooting Line interface class
//////////////////////////////////////////////////////////////////

// ShootingLine.h : Declaration of the CShootingLine

#ifndef __SHOOTINGLINE_H_
#define __SHOOTINGLINE_H_

#include "resource.h"       // main symbols
#include "ValidateShootingLine.h"

class DShootingRate;
class DMaxShootingRate;

/////////////////////////////////////////////////////////////////////////////
// CShootingLine Shooting line interface class
// =DOC SECTION Class ATL_NO_VTABLE
// =DOC SECTION 
// =DOC TEXT class ATL_NO_VTABLE extends nothing
// =DOC TEXT Shooting line interface class
//////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE CShootingLine : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CShootingLine, &CLSID_ShootingLine>,
	public ISupportErrorInfo,
	public IDispatchImpl<IShootingLine, &IID_IShootingLine, &LIBID_DPCALCLib>
{
public:

	// =DOC SECTION Method CShootingLine(): constructor
	// =DOC SECTION Class constructor
	// =DOC TEXT public CShootingLine()
	// =DOC TEXT Class constructor
	//////////////////////////////////////////////////////////////////
   CShootingLine();
   
	// =DOC SECTION Method ~CShootingLine(): destructor
	// =DOC SECTION Class destructor
	// =DOC TEXT public virtual ~CShootingLine()
	// =DOC TEXT Class destructor
	//////////////////////////////////////////////////////////////////
   virtual ~CShootingLine();

DECLARE_REGISTRY_RESOURCEID(IDR_SHOOTINGLINE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CShootingLine)
	COM_INTERFACE_ENTRY(IShootingLine)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IShootingLine
public:
	// =DOC SECTION Method GetMaxShootingRate( /*[in]*/ double FanCapacity, /*[in]*/ int RelativeHumidity, /*[out]*/ double* MaxShootingRate, /*[out]*/ BSTR* ErrMsg): STDMETHOD
	// =DOC SECTION Determine the maximum allowed shooting rate for a shooting line
	// =DOC TEXT public GetMaxShootingRate( /*[in]*/ double FanCapacity, /*[in]*/ int RelativeHumidity, /*[out]*/ double* MaxShootingRate, /*[out]*/ BSTR* ErrMsg): STDMETHOD
	// =DOC TEXT /*[in]*/ double FanCapacity: fan capacity
	// =DOC TEXT /*[in]*/ int RelativeHumidity: relative humidity
	// =DOC TEXT /*[out]*/ double* MaxShootingRate: calculated maximum shooting rate
	// =DOC TEXT /*[out]*/ BSTR* ErrMsg: error message if unsuccessful
	// =DOC TEXT Returns HRESULT: success (S_OK) or failure (E_FAIL)
	// =DOC TEXT Determine the maximum allowed shooting rate for a shooting line
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
	STDMETHOD(GetMaxShootingRate)( /*[in]*/ double FanCapacity,
                                  /*[in]*/ int RelativeHumidity,
                                  /*[out]*/ double* MaxShootingRate,
                                  /*[out]*/ BSTR* ErrMsg );

   // =DOC SECTION Method GetShootingRate(/*[in]*/ double CylinderTemp, /*[in]*/ double HoseLength, /*[in]*/ double HoseDiameter, /*[out]*/ double* ShootingRate, /*[out]*/ BSTR* ErrMsg ): STDMETHOD
	// =DOC SECTION Determine the shooting rate for a shooting line
	// =DOC TEXT public GetShootingRate(/*[in]*/ double CylinderTemp, /*[in]*/ double HoseLength, /*[in]*/ double HoseDiameter, /*[out]*/ double* ShootingRate, /*[out]*/ BSTR* ErrMsg ): STDMETHOD
	// =DOC TEXT /*[in]*/ double CylinderTemp: cylinder temperature
	// =DOC TEXT /*[in]*/ double HoseLength: hose length
	// =DOC TEXT /*[in]*/ double HoseDiameter: hose inside diameter
	// =DOC TEXT /*[out]*/ double* ShootingRate: return value
	// =DOC TEXT /*[out]*/ ErrMsg: error message if unsuccessful
	// =DOC TEXT Returns HRESULT: success (S_OK) or failure (E_FAIL)
	// =DOC TEXT Determine the shooting rate for a shooting line
   // =DOC TEXT Return S_OK if there were no problems and S_FALSE if there were problems during the calculation.
	//////////////////////////////////////////////////////////////////
   STDMETHOD(GetShootingRate)( /*[in]*/ double CylinderTemp,
                               /*[in]*/ double OneFourthLength,
                               /*[in]*/ double OneEightLength,
							   /*[in]*/	 double NumCylinders,
							   /*[in]*/  double PressureChosen,
                               /*[out]*/ double* ShootingRate,
							   /*[out]*/ double* TimeToEmptyCylinder,
                               /*[out]*/ BSTR* ErrMsg );

protected:
   // =DOC SECTION Member m_pMaxShootingRate: DMaxShootingRate*
	// =DOC SECTION Pointer to a Max Shooting Rate object
	// =DOC TEXT m_pMaxShootingRate: DMaxShootingRate*
	// =DOC TEXT Pointer to a Max Shooting Rate object
	//////////////////////////////////////////////////////////////////
	DMaxShootingRate* m_pMaxShootingRate;

   // =DOC SECTION Member m_pShootingRate: DShootingRate*
	// =DOC SECTION Pointer to a shooting rate object
	// =DOC TEXT m_pShootingRate: DShootingRate*
	// =DOC TEXT Pointer to a shooting rate object
	//////////////////////////////////////////////////////////////////
	DShootingRate* m_pShootingRate;

	// =DOC SECTION Member m_pValidator: DValidateShootingLine*
	// =DOC SECTION Pointer to a shooting line validation object
	// =DOC TEXT m_pValidator: DValidateShootingLine*
	// =DOC TEXT Pointer to a shooting line validation object
	//////////////////////////////////////////////////////////////////
	DValidateShootingLine* m_pValidator;

};

#endif //__SHOOTINGLINE_H_
