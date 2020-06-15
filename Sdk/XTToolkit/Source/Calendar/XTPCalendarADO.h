// Created by Microsoft (R) C/C++ Compiler Version 11.00.0000 (c8fff05a).
//
// This file is a part of the XTREME CALENDAR MFC class library.
// (c)1998-2011 Codejock Software, All Rights Reserved.
//
// THIS SOURCE FILE IS THE PROPERTY OF CODEJOCK SOFTWARE AND IS NOT TO BE
// RE-DISTRIBUTED BY ANY MEANS WHATSOEVER WITHOUT THE EXPRESSED WRITTEN
// CONSENT OF CODEJOCK SOFTWARE.
//
// THIS SOURCE CODE CAN ONLY BE USED UNDER THE TERMS AND CONDITIONS OUTLINED
// IN THE XTREME TOOLKIT PRO LICENSE AGREEMENT. CODEJOCK SOFTWARE GRANTS TO
// YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE THIS SOFTWARE ON A
// SINGLE COMPUTER.
//
// CONTACT INFORMATION:
// support@codejock.com
// http://www.codejock.com
//
/////////////////////////////////////////////////////////////////////////////

//{{AFX_CODEJOCK_PRIVATE
#if !defined(_XTPCALENDARADO_H__)
#define _XTPCALENDARADO_H__
//}}AFX_CODEJOCK_PRIVATE

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#pragma pack(push, 8)

#if (_MSC_VER <= 1100)
#pragma warning(disable:4510 4513 4610 4310 4244)
#endif

#include <comdef.h>

//{{AFX_CODEJOCK_PRIVATE
namespace XTPADOX {

//
// Forward references and typedefs
//

struct __declspec(uuid("00000512-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Collection;
struct __declspec(uuid("00000513-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _DynaCollection;
struct __declspec(uuid("00000603-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Catalog;
struct __declspec(uuid("00000611-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Tables;
struct /* coclass */ Table;
struct __declspec(uuid("00000610-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Table;
struct __declspec(uuid("0000061d-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Columns;
struct /* coclass */ Column;
struct __declspec(uuid("0000061c-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Column;
struct __declspec(uuid("00000504-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Properties;
struct __declspec(uuid("00000503-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Property;
struct __declspec(uuid("00000620-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Indexes;
struct /* coclass */ Index;
struct __declspec(uuid("0000061f-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Index;
struct __declspec(uuid("00000623-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Keys;
struct /* coclass */ Key;
struct __declspec(uuid("00000622-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Key;
struct __declspec(uuid("00000626-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Procedures;
struct __declspec(uuid("00000625-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Procedure;
struct __declspec(uuid("00000614-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Views;
struct __declspec(uuid("00000613-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ View;
struct __declspec(uuid("00000617-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Groups;
struct /* coclass */ Group;
struct __declspec(uuid("00000628-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Group;
struct __declspec(uuid("00000616-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Group25;
struct __declspec(uuid("0000061a-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Users;
struct /* coclass */ User;
struct __declspec(uuid("00000627-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _User;
struct __declspec(uuid("00000619-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _User25;
struct /* coclass */ Catalog;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(_Collection, __uuidof(_Collection));
_COM_SMARTPTR_TYPEDEF(_DynaCollection, __uuidof(_DynaCollection));
_COM_SMARTPTR_TYPEDEF(_Catalog, __uuidof(_Catalog));
_COM_SMARTPTR_TYPEDEF(Tables, __uuidof(Tables));
_COM_SMARTPTR_TYPEDEF(_Table, __uuidof(_Table));
_COM_SMARTPTR_TYPEDEF(Columns, __uuidof(Columns));
_COM_SMARTPTR_TYPEDEF(_Column, __uuidof(_Column));
_COM_SMARTPTR_TYPEDEF(Properties, __uuidof(Properties));
_COM_SMARTPTR_TYPEDEF(Property, __uuidof(Property));
_COM_SMARTPTR_TYPEDEF(Indexes, __uuidof(Indexes));
_COM_SMARTPTR_TYPEDEF(_Index, __uuidof(_Index));
_COM_SMARTPTR_TYPEDEF(Keys, __uuidof(Keys));
_COM_SMARTPTR_TYPEDEF(_Key, __uuidof(_Key));
_COM_SMARTPTR_TYPEDEF(Procedures, __uuidof(Procedures));
_COM_SMARTPTR_TYPEDEF(Procedure, __uuidof(Procedure));
_COM_SMARTPTR_TYPEDEF(Views, __uuidof(Views));
_COM_SMARTPTR_TYPEDEF(View, __uuidof(View));
_COM_SMARTPTR_TYPEDEF(Groups, __uuidof(Groups));
_COM_SMARTPTR_TYPEDEF(_Group, __uuidof(_Group));
_COM_SMARTPTR_TYPEDEF(_Group25, __uuidof(_Group25));
_COM_SMARTPTR_TYPEDEF(Users, __uuidof(Users));
_COM_SMARTPTR_TYPEDEF(_User, __uuidof(_User));
_COM_SMARTPTR_TYPEDEF(_User25, __uuidof(_User25));

//
// Type library items
//

enum AllowNullsEnum
{
	adIndexNullsAllow = 0,
	adIndexNullsDisallow = 1,
	adIndexNullsIgnore = 2,
	adIndexNullsIgnoreAny = 4
};

enum RuleEnum
{
	adRINone = 0,
	adRICascade = 1,
	adRISetNull = 2,
	adRISetDefault = 3
};

enum KeyTypeEnum
{
	adKeyPrimary = 1,
	adKeyForeign = 2,
	adKeyUnique = 3
};


enum ObjectTypeEnum
{
	adPermObjProviderSpecific = -1,
	adPermObjTable = 1,
	adPermObjColumn = 2,
	adPermObjDatabase = 3,
	adPermObjProcedure = 4,
	adPermObjView = 5
};

enum RightsEnum
{
	adRightNone = 0,
	adRightDrop = 256,
	adRightExclusive = 512,
	adRightReadDesign = 1024,
	adRightWriteDesign = 2048,
	adRightWithGrant = 4096,
	adRightReference = 8192,
	adRightCreate = 16384,
	adRightInsert = 32768,
	adRightDelete = 65536,
	adRightReadPermissions = 131072,
	adRightWritePermissions = 262144,
	adRightWriteOwner = 524288,
	adRightMaximumAllowed = 33554432,
	adRightFull = 268435456,
	adRightExecute = 536870912,
	adRightUpdate = 1073741824
};

enum ActionEnum
{
	adAccessGrant = 1,
	adAccessSet = 2,
	adAccessDeny = 3,
	adAccessRevoke = 4
};

enum InheritTypeEnum
{
	adInheritNone = 0,
	adInheritObjects = 1,
	adInheritContainers = 2,
	adInheritBoth = 3,
	adInheritNoPropogate = 4
};

enum ColumnAttributesEnum
{
	adColFixed = 1,
	adColNullable = 2
};

enum SortOrderEnum
{
	adSortAscending = 1,
	adSortDescending = 2
};

enum DataTypeEnumAdoX
{
	adEmpty = 0,
	adTinyInt = 16,
	adSmallInt = 2,
	adInteger = 3,
	adBigInt = 20,
	adUnsignedTinyInt = 17,
	adUnsignedSmallInt = 18,
	adUnsignedInt = 19,
	adUnsignedBigInt = 21,
	adSingle = 4,
	adDouble = 5,
	adCurrency = 6,
	adDecimal = 14,
	adNumeric = 131,
	adBoolean = 11,
	adError = 10,
	adUserDefined = 132,
	adVariant = 12,
	adIDispatch = 9,
	adIUnknown = 13,
	adGUID = 72,
	adDate = 7,
	adDBDate = 133,
	adDBTime = 134,
	adDBTimeStamp = 135,
	adBSTR = 8,
	adChar = 129,
	adVarChar = 200,
	adLongVarChar = 201,
	adWChar = 130,
	adVarWChar = 202,
	adLongVarWChar = 203,
	adBinary = 128,
	adVarBinary = 204,
	adLongVarBinary = 205,
	adChapter = 136,
	adFileTime = 64,
	adPropVariant = 138,
	adVarNumeric = 139
};


struct __declspec(uuid("00000512-0000-0010-8000-00aa006d2ea4"))
_Collection : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetCount))
	long Count;

	//
	// Wrapper methods for error-handling
	//

	long GetCount ();
	IUnknown * _NewEnum ();
	HRESULT Refresh ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Count (
		long * c) = 0;
	virtual HRESULT __stdcall raw__NewEnum (
		IUnknown * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Refresh () = 0;
};

struct __declspec(uuid("00000513-0000-0010-8000-00aa006d2ea4"))
_DynaCollection : public _Collection
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT Append (
		IDispatch * Object);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Append (
		IDispatch * Object) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000603-0000-0010-8000-00aa006d2ea4"))
_Catalog : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetTables))
	TablesPtr Tables;
	__declspec(property(get = GetProcedures))
	ProceduresPtr Procedures;
	__declspec(property(get = GetViews))
	ViewsPtr Views;
	__declspec(property(get = GetGroups))
	GroupsPtr Groups;
	__declspec(property(get = GetUsers))
	UsersPtr Users;

	//
	// Wrapper methods for error-handling
	//

	TablesPtr GetTables ();
	_variant_t GetActiveConnection ();
	void PutActiveConnection (
		const _variant_t & pVal);
	void PutRefActiveConnection (
		IDispatch * pVal);
	ProceduresPtr GetProcedures ();
	ViewsPtr GetViews ();
	GroupsPtr GetGroups ();
	UsersPtr GetUsers ();
	_variant_t Create (
		_bstr_t ConnectString);
	_bstr_t GetObjectOwner (
		_bstr_t ObjectName,
		enum ObjectTypeEnum ObjectType,
		const _variant_t & ObjectTypeId = vtMissing);
	HRESULT SetObjectOwner (
		_bstr_t ObjectName,
		enum ObjectTypeEnum ObjectType,
		_bstr_t UserName,
		const _variant_t & ObjectTypeId = vtMissing);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Tables (
		struct Tables * * ppvObject) = 0;
	virtual HRESULT __stdcall get_ActiveConnection (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall put_ActiveConnection (
		VARIANT pVal) = 0;
	virtual HRESULT __stdcall putref_ActiveConnection (
		IDispatch * pVal) = 0;
	virtual HRESULT __stdcall get_Procedures (
		struct Procedures * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Views (
		struct Views * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Groups (
		struct Groups * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Users (
		struct Users * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Create (
		BSTR ConnectString,
		VARIANT * Connection) = 0;
	virtual HRESULT __stdcall raw_GetObjectOwner (
		BSTR ObjectName,
		enum ObjectTypeEnum ObjectType,
		VARIANT ObjectTypeId,
		BSTR * OwnerName) = 0;
	virtual HRESULT __stdcall raw_SetObjectOwner (
		BSTR ObjectName,
		enum ObjectTypeEnum ObjectType,
		BSTR UserName,
		VARIANT ObjectTypeId = vtMissing) = 0;
};

struct __declspec(uuid("00000611-0000-0010-8000-00aa006d2ea4"))
Tables : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_TablePtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_TablePtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _Table * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000609-0000-0010-8000-00aa006d2ea4"))
Table;
	// [ default ] interface _Table

struct __declspec(uuid("00000610-0000-0010-8000-00aa006d2ea4"))
_Table : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetColumns))
	ColumnsPtr Columns;
	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetType))
	_bstr_t Type;
	__declspec(property(get = GetIndexes))
	IndexesPtr Indexes;
	__declspec(property(get = GetKeys))
	KeysPtr Keys;
	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;
	__declspec(property(get = GetDateCreated))
	_variant_t DateCreated;
	__declspec(property(get = GetDateModified))
	_variant_t DateModified;
	__declspec(property(get = GetParentCatalog, put = PutRefParentCatalog))
	_CatalogPtr ParentCatalog;

	//
	// Wrapper methods for error-handling
	//

	ColumnsPtr GetColumns ();
	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	_bstr_t GetType ();
	IndexesPtr GetIndexes ();
	KeysPtr GetKeys ();
	PropertiesPtr GetProperties ();
	_variant_t GetDateCreated ();
	_variant_t GetDateModified ();
	_CatalogPtr GetParentCatalog ();
	void PutParentCatalog (
		struct _Catalog * ppvObject);
	void PutRefParentCatalog (
		struct _Catalog * ppvObject);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Columns (
		struct Columns * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_Type (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall get_Indexes (
		struct Indexes * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Keys (
		struct Keys * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
	virtual HRESULT __stdcall get_DateCreated (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall get_DateModified (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall get_ParentCatalog (
		struct _Catalog * * ppvObject) = 0;
	virtual HRESULT __stdcall put_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
	virtual HRESULT __stdcall putref_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
};

struct __declspec(uuid("0000061d-0000-0010-8000-00aa006d2ea4"))
Columns : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_ColumnPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_ColumnPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item,
		enum DataTypeEnumAdoX Type,
		long DefinedSize);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _Column * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item,
		enum DataTypeEnumAdoX Type,
		long DefinedSize) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("0000061b-0000-0010-8000-00aa006d2ea4"))
Column;
	// [ default ] interface _Column

struct __declspec(uuid("0000061c-0000-0010-8000-00aa006d2ea4"))
_Column : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	enum ColumnAttributesEnum Attributes;
	__declspec(property(get = GetDefinedSize, put = PutDefinedSize))
	long DefinedSize;
	__declspec(property(get = GetNumericScale, put = PutNumericScale))
	unsigned char NumericScale;
	__declspec(property(get = GetPrecision, put = PutPrecision))
	long Precision;
	__declspec(property(get = GetRelatedColumn, put = PutRelatedColumn))
	_bstr_t RelatedColumn;
	__declspec(property(get = GetSortOrder, put = PutSortOrder))
	enum SortOrderEnum SortOrder;
	__declspec(property(get = GetType, put = PutType))
	enum DataTypeEnumAdoX Type;
	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;
	__declspec(property(get = GetParentCatalog, put = PutRefParentCatalog))
	_CatalogPtr ParentCatalog;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	enum ColumnAttributesEnum GetAttributes ();
	void PutAttributes (
		enum ColumnAttributesEnum pVal);
	long GetDefinedSize ();
	void PutDefinedSize (
		long pVal);
	unsigned char GetNumericScale ();
	void PutNumericScale (
		unsigned char pVal);
	long GetPrecision ();
	void PutPrecision (
		long pVal);
	_bstr_t GetRelatedColumn ();
	void PutRelatedColumn (
		_bstr_t pVal);
	enum SortOrderEnum GetSortOrder ();
	void PutSortOrder (
		enum SortOrderEnum pVal);
	enum DataTypeEnumAdoX GetType ();
	void PutType (
		enum DataTypeEnumAdoX pVal);
	PropertiesPtr GetProperties ();
	_CatalogPtr GetParentCatalog ();
	void PutParentCatalog (
		struct _Catalog * ppvObject);
	void PutRefParentCatalog (
		struct _Catalog * ppvObject);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_Attributes (
		enum ColumnAttributesEnum * pVal) = 0;
	virtual HRESULT __stdcall put_Attributes (
		enum ColumnAttributesEnum pVal) = 0;
	virtual HRESULT __stdcall get_DefinedSize (
		long * pVal) = 0;
	virtual HRESULT __stdcall put_DefinedSize (
		long pVal) = 0;
	virtual HRESULT __stdcall get_NumericScale (
		unsigned char * pVal) = 0;
	virtual HRESULT __stdcall put_NumericScale (
		unsigned char pVal) = 0;
	virtual HRESULT __stdcall get_Precision (
		long * pVal) = 0;
	virtual HRESULT __stdcall put_Precision (
		long pVal) = 0;
	virtual HRESULT __stdcall get_RelatedColumn (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_RelatedColumn (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_SortOrder (
		enum SortOrderEnum * pVal) = 0;
	virtual HRESULT __stdcall put_SortOrder (
		enum SortOrderEnum pVal) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnumAdoX * pVal) = 0;
	virtual HRESULT __stdcall put_Type (
		enum DataTypeEnumAdoX pVal) = 0;
	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
	virtual HRESULT __stdcall get_ParentCatalog (
		struct _Catalog * * ppvObject) = 0;
	virtual HRESULT __stdcall put_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
	virtual HRESULT __stdcall putref_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
};

struct __declspec(uuid("00000504-0000-0010-8000-00aa006d2ea4"))
Properties : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	PropertyPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	PropertyPtr GetItem (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct Property * * ppvObject) = 0;
};

struct __declspec(uuid("00000503-0000-0010-8000-00aa006d2ea4"))
Property : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetValue, put = PutValue))
	_variant_t Value;
	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetType))
	enum DataTypeEnumAdoX Type;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	long Attributes;

	//
	// Wrapper methods for error-handling
	//

	_variant_t GetValue ();
	void PutValue (
		const _variant_t & pVal);
	_bstr_t GetName ();
	enum DataTypeEnumAdoX GetType ();
	long GetAttributes ();
	void PutAttributes (
		long plAttributes);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Value (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall put_Value (
		VARIANT pVal) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnumAdoX * ptype) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * plAttributes) = 0;
	virtual HRESULT __stdcall put_Attributes (
		long plAttributes) = 0;
};

struct __declspec(uuid("00000620-0000-0010-8000-00aa006d2ea4"))
Indexes : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_IndexPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_IndexPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item,
		const _variant_t & Columns = vtMissing);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _Index * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item,
		VARIANT Columns = vtMissing) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("0000061e-0000-0010-8000-00aa006d2ea4"))
Index;
	// [ default ] interface _Index

struct __declspec(uuid("0000061f-0000-0010-8000-00aa006d2ea4"))
_Index : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetClustered, put = PutClustered))
	VARIANT_BOOL Clustered;
	__declspec(property(get = GetIndexNulls, put = PutIndexNulls))
	enum AllowNullsEnum IndexNulls;
	__declspec(property(get = GetPrimaryKey, put = PutPrimaryKey))
	VARIANT_BOOL PrimaryKey;
	__declspec(property(get = GetUnique, put = PutUnique))
	VARIANT_BOOL Unique;
	__declspec(property(get = GetColumns))
	ColumnsPtr Columns;
	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	VARIANT_BOOL GetClustered ();
	void PutClustered (
		VARIANT_BOOL pVal);
	enum AllowNullsEnum GetIndexNulls ();
	void PutIndexNulls (
		enum AllowNullsEnum pVal);
	VARIANT_BOOL GetPrimaryKey ();
	void PutPrimaryKey (
		VARIANT_BOOL pVal);
	VARIANT_BOOL GetUnique ();
	void PutUnique (
		VARIANT_BOOL pVal);
	ColumnsPtr GetColumns ();
	PropertiesPtr GetProperties ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_Clustered (
		VARIANT_BOOL * pVal) = 0;
	virtual HRESULT __stdcall put_Clustered (
		VARIANT_BOOL pVal) = 0;
	virtual HRESULT __stdcall get_IndexNulls (
		enum AllowNullsEnum * pVal) = 0;
	virtual HRESULT __stdcall put_IndexNulls (
		enum AllowNullsEnum pVal) = 0;
	virtual HRESULT __stdcall get_PrimaryKey (
		VARIANT_BOOL * pVal) = 0;
	virtual HRESULT __stdcall put_PrimaryKey (
		VARIANT_BOOL pVal) = 0;
	virtual HRESULT __stdcall get_Unique (
		VARIANT_BOOL * pVal) = 0;
	virtual HRESULT __stdcall put_Unique (
		VARIANT_BOOL pVal) = 0;
	virtual HRESULT __stdcall get_Columns (
		struct Columns * * ppvObject) = 0;
	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
};

struct __declspec(uuid("00000623-0000-0010-8000-00aa006d2ea4"))
Keys : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_KeyPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_KeyPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item,
		enum KeyTypeEnum Type,
		const _variant_t & Column,
		_bstr_t RelatedTable,
		_bstr_t RelatedColumn);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _Key * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item,
		enum KeyTypeEnum Type,
		VARIANT Column,
		BSTR RelatedTable,
		BSTR RelatedColumn) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000621-0000-0010-8000-00aa006d2ea4"))
Key;
	// [ default ] interface _Key

struct __declspec(uuid("00000622-0000-0010-8000-00aa006d2ea4"))
_Key : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetDeleteRule, put = PutDeleteRule))
	enum RuleEnum DeleteRule;
	__declspec(property(get = GetType, put = PutType))
	enum KeyTypeEnum Type;
	__declspec(property(get = GetRelatedTable, put = PutRelatedTable))
	_bstr_t RelatedTable;
	__declspec(property(get = GetUpdateRule, put = PutUpdateRule))
	enum RuleEnum UpdateRule;
	__declspec(property(get = GetColumns))
	ColumnsPtr Columns;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	enum RuleEnum GetDeleteRule ();
	void PutDeleteRule (
		enum RuleEnum pVal);
	enum KeyTypeEnum GetType ();
	void PutType (
		enum KeyTypeEnum pVal);
	_bstr_t GetRelatedTable ();
	void PutRelatedTable (
		_bstr_t pVal);
	enum RuleEnum GetUpdateRule ();
	void PutUpdateRule (
		enum RuleEnum pVal);
	ColumnsPtr GetColumns ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_DeleteRule (
		enum RuleEnum * pVal) = 0;
	virtual HRESULT __stdcall put_DeleteRule (
		enum RuleEnum pVal) = 0;
	virtual HRESULT __stdcall get_Type (
		enum KeyTypeEnum * pVal) = 0;
	virtual HRESULT __stdcall put_Type (
		enum KeyTypeEnum pVal) = 0;
	virtual HRESULT __stdcall get_RelatedTable (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_RelatedTable (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall get_UpdateRule (
		enum RuleEnum * pVal) = 0;
	virtual HRESULT __stdcall put_UpdateRule (
		enum RuleEnum pVal) = 0;
	virtual HRESULT __stdcall get_Columns (
		struct Columns * * ppvObject) = 0;
};

struct __declspec(uuid("00000626-0000-0010-8000-00aa006d2ea4"))
Procedures : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	ProcedurePtr Item[];

	//
	// Wrapper methods for error-handling
	//

	ProcedurePtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		_bstr_t Name,
		IDispatch * Command);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct Procedure * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		BSTR Name,
		IDispatch * Command) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000625-0000-0010-8000-00aa006d2ea4"))
Procedure : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetDateCreated))
	_variant_t DateCreated;
	__declspec(property(get = GetDateModified))
	_variant_t DateModified;

	//
	// Wrapper methods for error-handling
	//

	_variant_t GetCommand ();
	void PutCommand (
		const _variant_t & pVar);
	void PutRefCommand (
		IDispatch * pVar);
	_bstr_t GetName ();
	_variant_t GetDateCreated ();
	_variant_t GetDateModified ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Command (
		VARIANT * pVar) = 0;
	virtual HRESULT __stdcall put_Command (
		VARIANT pVar) = 0;
	virtual HRESULT __stdcall putref_Command (
		IDispatch * pVar) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall get_DateCreated (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall get_DateModified (
		VARIANT * pVal) = 0;
};

struct __declspec(uuid("00000614-0000-0010-8000-00aa006d2ea4"))
Views : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	ViewPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	ViewPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		_bstr_t Name,
		IDispatch * Command);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct View * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		BSTR Name,
		IDispatch * Command) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000613-0000-0010-8000-00aa006d2ea4"))
View : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetDateCreated))
	_variant_t DateCreated;
	__declspec(property(get = GetDateModified))
	_variant_t DateModified;

	//
	// Wrapper methods for error-handling
	//

	_variant_t GetCommand ();
	void PutCommand (
		const _variant_t & pVal);
	void PutRefCommand (
		IDispatch * pVal);
	_bstr_t GetName ();
	_variant_t GetDateCreated ();
	_variant_t GetDateModified ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Command (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall put_Command (
		VARIANT pVal) = 0;
	virtual HRESULT __stdcall putref_Command (
		IDispatch * pVal) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall get_DateCreated (
		VARIANT * pVal) = 0;
	virtual HRESULT __stdcall get_DateModified (
		VARIANT * pVal) = 0;
};

struct __declspec(uuid("00000617-0000-0010-8000-00aa006d2ea4"))
Groups : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_GroupPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_GroupPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _Group * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000615-0000-0010-8000-00aa006d2ea4"))
Group;
	// [ default ] interface _Group


struct __declspec(uuid("00000616-0000-0010-8000-00aa006d2ea4"))
_Group25 : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetUsers))
	UsersPtr Users;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	enum RightsEnum GetPermissions (
		const _variant_t & Name,
		enum ObjectTypeEnum ObjectType,
		const _variant_t & ObjectTypeId = vtMissing);
	HRESULT SetPermissions (
		const _variant_t & Name,
		enum ObjectTypeEnum ObjectType,
		enum ActionEnum Action,
		enum RightsEnum Rights,
		enum InheritTypeEnum Inherit,
		const _variant_t & ObjectTypeId = vtMissing);
	UsersPtr GetUsers ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall raw_GetPermissions (
		VARIANT Name,
		enum ObjectTypeEnum ObjectType,
		VARIANT ObjectTypeId,
		enum RightsEnum * Rights) = 0;
	virtual HRESULT __stdcall raw_SetPermissions (
		VARIANT Name,
		enum ObjectTypeEnum ObjectType,
		enum ActionEnum Action,
		enum RightsEnum Rights,
		enum InheritTypeEnum Inherit,
		VARIANT ObjectTypeId = vtMissing) = 0;
	virtual HRESULT __stdcall get_Users (
		struct Users * * ppvObject) = 0;
};

struct __declspec(uuid("00000628-0000-0010-8000-00aa006d2ea4"))
_Group : public _Group25
{
	//
	// Property data
	//

	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;
	__declspec(property(get = GetParentCatalog, put = PutRefParentCatalog))
	_CatalogPtr ParentCatalog;

	//
	// Wrapper methods for error-handling
	//

	PropertiesPtr GetProperties ();
	_CatalogPtr GetParentCatalog ();
	void PutParentCatalog (
		struct _Catalog * ppvObject);
	void PutRefParentCatalog (
		struct _Catalog * ppvObject);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
	virtual HRESULT __stdcall get_ParentCatalog (
		struct _Catalog * * ppvObject) = 0;
	virtual HRESULT __stdcall put_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
	virtual HRESULT __stdcall putref_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
};

struct __declspec(uuid("0000061a-0000-0010-8000-00aa006d2ea4"))
Users : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_UserPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_UserPtr GetItem (
		const _variant_t & Item);
	HRESULT Append (
		const _variant_t & Item,
		_bstr_t Password);
	HRESULT Delete (
		const _variant_t & Item);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Item,
		struct _User * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Append (
		VARIANT Item,
		BSTR Password) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Item) = 0;
};

struct __declspec(uuid("00000618-0000-0010-8000-00aa006d2ea4"))
User;
	// [ default ] interface _User


struct __declspec(uuid("00000619-0000-0010-8000-00aa006d2ea4"))
_User25 : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetGroups))
	GroupsPtr Groups;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pVal);
	enum RightsEnum GetPermissions (
		const _variant_t & Name,
		enum ObjectTypeEnum ObjectType,
		const _variant_t & ObjectTypeId = vtMissing);
	HRESULT SetPermissions (
		const _variant_t & Name,
		enum ObjectTypeEnum ObjectType,
		enum ActionEnum Action,
		enum RightsEnum Rights,
		enum InheritTypeEnum Inherit,
		const _variant_t & ObjectTypeId = vtMissing);
	HRESULT ChangePassword (
		_bstr_t OldPassword,
		_bstr_t NewPassword);
	GroupsPtr GetGroups ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pVal) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pVal) = 0;
	virtual HRESULT __stdcall raw_GetPermissions (
		VARIANT Name,
		enum ObjectTypeEnum ObjectType,
		VARIANT ObjectTypeId,
		enum RightsEnum * Rights) = 0;
	virtual HRESULT __stdcall raw_SetPermissions (
		VARIANT Name,
		enum ObjectTypeEnum ObjectType,
		enum ActionEnum Action,
		enum RightsEnum Rights,
		enum InheritTypeEnum Inherit,
		VARIANT ObjectTypeId = vtMissing) = 0;
	virtual HRESULT __stdcall raw_ChangePassword (
		BSTR OldPassword,
		BSTR NewPassword) = 0;
	virtual HRESULT __stdcall get_Groups (
		struct Groups * * ppvObject) = 0;
};

struct __declspec(uuid("00000627-0000-0010-8000-00aa006d2ea4"))
_User : public _User25
{
	//
	// Property data
	//

	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;
	__declspec(property(get = GetParentCatalog, put = PutRefParentCatalog))
	_CatalogPtr ParentCatalog;

	//
	// Wrapper methods for error-handling
	//

	PropertiesPtr GetProperties ();
	_CatalogPtr GetParentCatalog ();
	void PutParentCatalog (
		struct _Catalog * ppvObject);
	void PutRefParentCatalog (
		struct _Catalog * ppvObject);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
	virtual HRESULT __stdcall get_ParentCatalog (
		struct _Catalog * * ppvObject) = 0;
	virtual HRESULT __stdcall put_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
	virtual HRESULT __stdcall putref_ParentCatalog (
		struct _Catalog * ppvObject) = 0;
};


struct __declspec(uuid("00000602-0000-0010-8000-00aa006d2ea4"))
Catalog;
	// [ default ] interface _Catalog

//
// Wrapper method implementations
//

//#include "Debug/msadox.tli"

} // namespace XTPADOX


namespace XTPADODB {

//
// Forward references and typedefs
//

typedef enum PositionEnum PositionEnum_Param;
typedef enum SearchDirectionEnum SearchDirection;
struct __declspec(uuid("00000512-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Collection;
struct __declspec(uuid("00000513-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _DynaCollection;
struct __declspec(uuid("00000534-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _ADO;
struct __declspec(uuid("00000504-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Properties;
struct __declspec(uuid("00000503-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Property;
struct __declspec(uuid("00000500-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Error;
struct __declspec(uuid("00000501-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Errors;
struct __declspec(uuid("00000508-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Command15;
struct __declspec(uuid("00000550-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Connection;
struct __declspec(uuid("00000515-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Connection15;
struct __declspec(uuid("00000556-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Recordset;
struct __declspec(uuid("00000555-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Recordset21;
struct __declspec(uuid("0000054f-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Recordset20;
struct __declspec(uuid("0000050e-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Recordset15;
struct __declspec(uuid("00000564-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Fields;
struct __declspec(uuid("0000054d-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Fields20;
struct __declspec(uuid("00000506-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Fields15;
struct __declspec(uuid("00000569-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Field;
struct __declspec(uuid("0000054c-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Field20;
typedef long ADO_LONGPTR;
struct __declspec(uuid("0000050c-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Parameter;
struct __declspec(uuid("0000050d-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Parameters;
struct __declspec(uuid("0000054e-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Command25;
struct __declspec(uuid("b08400bd-f9d1-4d02-b856-71d5dba123e9"))
/* dual interface */ _Command;
struct __declspec(uuid("00000402-0000-0010-8000-00aa006d2ea4"))
/* interface */ ConnectionEventsVt;
struct __declspec(uuid("00000403-0000-0010-8000-00aa006d2ea4"))
/* interface */ RecordsetEventsVt;
struct __declspec(uuid("00000400-0000-0010-8000-00aa006d2ea4"))
/* dispinterface */ ConnectionEvents;
struct __declspec(uuid("00000266-0000-0010-8000-00aa006d2ea4"))
/* dispinterface */ RecordsetEvents;
struct __declspec(uuid("00000516-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADOConnectionConstruction15;
struct __declspec(uuid("00000551-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADOConnectionConstruction;
struct /* coclass */ Connection;
struct __declspec(uuid("00000562-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Record;
struct /* coclass */ Record;
struct __declspec(uuid("00000565-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ _Stream;
struct /* coclass */ Stream;
struct __declspec(uuid("00000567-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADORecordConstruction;
struct __declspec(uuid("00000568-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADOStreamConstruction;
struct __declspec(uuid("00000517-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADOCommandConstruction;
struct /* coclass */ Command;
struct /* coclass */ Recordset;
struct __declspec(uuid("00000283-0000-0010-8000-00aa006d2ea4"))
/* interface */ ADORecordsetConstruction;
struct __declspec(uuid("00000505-0000-0010-8000-00aa006d2ea4"))
/* dual interface */ Field15;
struct /* coclass */ Parameter;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(_Collection, __uuidof(_Collection));
_COM_SMARTPTR_TYPEDEF(_DynaCollection, __uuidof(_DynaCollection));
_COM_SMARTPTR_TYPEDEF(_ADO, __uuidof(_ADO));
_COM_SMARTPTR_TYPEDEF(Properties, __uuidof(Properties));
_COM_SMARTPTR_TYPEDEF(Property, __uuidof(Property));
_COM_SMARTPTR_TYPEDEF(Error, __uuidof(Error));
_COM_SMARTPTR_TYPEDEF(Errors, __uuidof(Errors));
_COM_SMARTPTR_TYPEDEF(Command15, __uuidof(Command15));
_COM_SMARTPTR_TYPEDEF(Connection15, __uuidof(Connection15));
_COM_SMARTPTR_TYPEDEF(_Connection, __uuidof(_Connection));
_COM_SMARTPTR_TYPEDEF(Recordset15, __uuidof(Recordset15));
_COM_SMARTPTR_TYPEDEF(Recordset20, __uuidof(Recordset20));
_COM_SMARTPTR_TYPEDEF(Recordset21, __uuidof(Recordset21));
_COM_SMARTPTR_TYPEDEF(_Recordset, __uuidof(_Recordset));
_COM_SMARTPTR_TYPEDEF(Fields15, __uuidof(Fields15));
_COM_SMARTPTR_TYPEDEF(Fields20, __uuidof(Fields20));
_COM_SMARTPTR_TYPEDEF(Fields, __uuidof(Fields));
_COM_SMARTPTR_TYPEDEF(Field20, __uuidof(Field20));
_COM_SMARTPTR_TYPEDEF(Field, __uuidof(Field));
_COM_SMARTPTR_TYPEDEF(_Parameter, __uuidof(_Parameter));
_COM_SMARTPTR_TYPEDEF(Parameters, __uuidof(Parameters));
_COM_SMARTPTR_TYPEDEF(Command25, __uuidof(Command25));
_COM_SMARTPTR_TYPEDEF(_Command, __uuidof(_Command));
_COM_SMARTPTR_TYPEDEF(ConnectionEventsVt, __uuidof(ConnectionEventsVt));
_COM_SMARTPTR_TYPEDEF(RecordsetEventsVt, __uuidof(RecordsetEventsVt));
_COM_SMARTPTR_TYPEDEF(ConnectionEvents, __uuidof(IDispatch));
_COM_SMARTPTR_TYPEDEF(RecordsetEvents, __uuidof(IDispatch));
_COM_SMARTPTR_TYPEDEF(ADOConnectionConstruction15, __uuidof(ADOConnectionConstruction15));
_COM_SMARTPTR_TYPEDEF(ADOConnectionConstruction, __uuidof(ADOConnectionConstruction));
_COM_SMARTPTR_TYPEDEF(_Record, __uuidof(_Record));
_COM_SMARTPTR_TYPEDEF(_Stream, __uuidof(_Stream));
_COM_SMARTPTR_TYPEDEF(ADORecordConstruction, __uuidof(ADORecordConstruction));
_COM_SMARTPTR_TYPEDEF(ADOStreamConstruction, __uuidof(ADOStreamConstruction));
_COM_SMARTPTR_TYPEDEF(ADOCommandConstruction, __uuidof(ADOCommandConstruction));
_COM_SMARTPTR_TYPEDEF(ADORecordsetConstruction, __uuidof(ADORecordsetConstruction));
_COM_SMARTPTR_TYPEDEF(Field15, __uuidof(Field15));

//
// Type library items
//

enum CursorTypeEnum
{
	adOpenUnspecified = -1,
	adOpenForwardOnly = 0,
	adOpenKeyset = 1,
	adOpenDynamic = 2,
	adOpenStatic = 3
};

enum CursorOptionEnum
{
	adHoldRecords = 256,
	adMovePrevious = 512,
	adAddNew = 16778240,
	adDelete = 16779264,
	adUpdate = 16809984,
	adBookmark = 8192,
	adApproxPosition = 16384,
	adUpdateBatch = 65536,
	adResync = 131072,
	adNotify = 262144,
	adFind = 524288,
	adSeek = 4194304,
	adIndex = 8388608
};

enum LockTypeEnum
{
	adLockUnspecified = -1,
	adLockReadOnly = 1,
	adLockPessimistic = 2,
	adLockOptimistic = 3,
	adLockBatchOptimistic = 4
};

enum ExecuteOptionEnum
{
	adOptionUnspecified = -1,
	adAsyncExecute = 16,
	adAsyncFetch = 32,
	adAsyncFetchNonBlocking = 64,
	adExecuteNoRecords = 128,
	adExecuteStream = 1024,
	adExecuteRecord = 2048
};

enum ConnectOptionEnum
{
	adConnectUnspecified = -1,
	adAsyncConnect = 16
};

enum ObjectStateEnum
{
	adStateClosed = 0,
	adStateOpen = 1,
	adStateConnecting = 2,
	adStateExecuting = 4,
	adStateFetching = 8
};

enum CursorLocationEnum
{
	adUseNone = 1,
	adUseServer = 2,
	adUseClient = 3,
	adUseClientBatch = 3
};

enum DataTypeEnum
{
	adEmpty = 0,
	adTinyInt = 16,
	adSmallInt = 2,
	adInteger = 3,
	adBigInt = 20,
	adUnsignedTinyInt = 17,
	adUnsignedSmallInt = 18,
	adUnsignedInt = 19,
	adUnsignedBigInt = 21,
	adSingle = 4,
	adDouble = 5,
	adCurrency = 6,
	adDecimal = 14,
	adNumeric = 131,
	adBoolean = 11,
	adError = 10,
	adUserDefined = 132,
	adVariant = 12,
	adIDispatch = 9,
	adIUnknown = 13,
	adGUID = 72,
	adDate = 7,
	adDBDate = 133,
	adDBTime = 134,
	adDBTimeStamp = 135,
	adBSTR = 8,
	adChar = 129,
	adVarChar = 200,
	adLongVarChar = 201,
	adWChar = 130,
	adVarWChar = 202,
	adLongVarWChar = 203,
	adBinary = 128,
	adVarBinary = 204,
	adLongVarBinary = 205,
	adChapter = 136,
	adFileTime = 64,
	adPropVariant = 138,
	adVarNumeric = 139,
	adArray = 8192
};

enum FieldAttributeEnum
{
	adFldUnspecified = -1,
	adFldMayDefer = 2,
	adFldUpdatable = 4,
	adFldUnknownUpdatable = 8,
	adFldFixed = 16,
	adFldIsNullable = 32,
	adFldMayBeNull = 64,
	adFldLong = 128,
	adFldRowID = 256,
	adFldRowVersion = 512,
	adFldCacheDeferred = 4096,
	adFldIsChapter = 8192,
	adFldNegativeScale = 16384,
	adFldKeyColumn = 32768,
	adFldIsRowURL = 65536,
	adFldIsDefaultStream = 131072,
	adFldIsCollection = 262144
};

enum EditModeEnum
{
	adEditNone = 0,
	adEditInProgress = 1,
	adEditAdd = 2,
	adEditDelete = 4
};

enum RecordStatusEnum
{
	adRecOK = 0,
	adRecNew = 1,
	adRecModified = 2,
	adRecDeleted = 4,
	adRecUnmodified = 8,
	adRecInvalid = 16,
	adRecMultipleChanges = 64,
	adRecPendingChanges = 128,
	adRecCanceled = 256,
	adRecCantRelease = 1024,
	adRecConcurrencyViolation = 2048,
	adRecIntegrityViolation = 4096,
	adRecMaxChangesExceeded = 8192,
	adRecObjectOpen = 16384,
	adRecOutOfMemory = 32768,
	adRecPermissionDenied = 65536,
	adRecSchemaViolation = 131072,
	adRecDBDeleted = 262144
};

enum GetRowsOptionEnum
{
	adGetRowsRest = -1
};

enum PositionEnum
{
	adPosUnknown = -1,
	adPosBOF = -2,
	adPosEOF = -3
};

enum BookmarkEnum
{
	adBookmarkCurrent = 0,
	adBookmarkFirst = 1,
	adBookmarkLast = 2
};

enum MarshalOptionsEnum
{
	adMarshalAll = 0,
	adMarshalModifiedOnly = 1
};

enum AffectEnum
{
	adAffectCurrent = 1,
	adAffectGroup = 2,
	adAffectAll = 3,
	adAffectAllChapters = 4
};

enum ResyncEnum
{
	adResyncUnderlyingValues = 1,
	adResyncAllValues = 2
};

enum CompareEnum
{
	adCompareLessThan = 0,
	adCompareEqual = 1,
	adCompareGreaterThan = 2,
	adCompareNotEqual = 3,
	adCompareNotComparable = 4
};

enum FilterGroupEnum
{
	adFilterNone = 0,
	adFilterPendingRecords = 1,
	adFilterAffectedRecords = 2,
	adFilterFetchedRecords = 3,
	adFilterPredicate = 4,
	adFilterConflictingRecords = 5
};

enum SearchDirectionEnum
{
	adSearchForward = 1,
	adSearchBackward = -1
};

enum PersistFormatEnum
{
	adPersistADTG = 0,
	adPersistXML = 1
};

enum StringFormatEnum
{
	adClipString = 2
};

enum ConnectPromptEnum
{
	adPromptAlways = 1,
	adPromptComplete = 2,
	adPromptCompleteRequired = 3,
	adPromptNever = 4
};

enum ConnectModeEnum
{
	adModeUnknown = 0,
	adModeRead = 1,
	adModeWrite = 2,
	adModeReadWrite = 3,
	adModeShareDenyRead = 4,
	adModeShareDenyWrite = 8,
	adModeShareExclusive = 12,
	adModeShareDenyNone = 16,
	adModeRecursive = 4194304
};

enum RecordCreateOptionsEnum
{
	adCreateCollection = 8192,
	adCreateNonCollection = 0,
	adOpenIfExists = 33554432,
	adCreateOverwrite = 67108864
};

enum RecordOpenOptionsEnum
{
	adOpenRecordUnspecified = -1,
	adOpenSource = 8388608,
	adOpenOutput = 8388608,
	adOpenAsync = 4096,
	adDelayFetchStream = 16384,
	adDelayFetchFields = 32768,
	adOpenExecuteCommand = 65536
};

enum IsolationLevelEnum
{
	adXactUnspecified = -1,
	adXactChaos = 16,
	adXactReadUncommitted = 256,
	adXactBrowse = 256,
	adXactCursorStability = 4096,
	adXactReadCommitted = 4096,
	adXactRepeatableRead = 65536,
	adXactSerializable = 1048576,
	adXactIsolated = 1048576
};

enum XactAttributeEnum
{
	adXactCommitRetaining = 131072,
	adXactAbortRetaining = 262144,
	adXactAsyncPhaseOne = 524288,
	adXactSyncPhaseOne = 1048576
};

enum PropertyAttributesEnum
{
	adPropNotSupported = 0,
	adPropRequired = 1,
	adPropOptional = 2,
	adPropRead = 512,
	adPropWrite = 1024
};

enum ErrorValueEnum
{
	adErrProviderFailed = 3000,
	adErrInvalidArgument = 3001,
	adErrOpeningFile = 3002,
	adErrReadFile = 3003,
	adErrWriteFile = 3004,
	adErrNoCurrentRecord = 3021,
	adErrIllegalOperation = 3219,
	adErrCantChangeProvider = 3220,
	adErrInTransaction = 3246,
	adErrFeatureNotAvailable = 3251,
	adErrItemNotFound = 3265,
	adErrObjectInCollection = 3367,
	adErrObjectNotSet = 3420,
	adErrDataConversion = 3421,
	adErrObjectClosed = 3704,
	adErrObjectOpen = 3705,
	adErrProviderNotFound = 3706,
	adErrBoundToCommand = 3707,
	adErrInvalidParamInfo = 3708,
	adErrInvalidConnection = 3709,
	adErrNotReentrant = 3710,
	adErrStillExecuting = 3711,
	adErrOperationCancelled = 3712,
	adErrStillConnecting = 3713,
	adErrInvalidTransaction = 3714,
	adErrNotExecuting = 3715,
	adErrUnsafeOperation = 3716,
	adwrnSecurityDialog = 3717,
	adwrnSecurityDialogHeader = 3718,
	adErrIntegrityViolation = 3719,
	adErrPermissionDenied = 3720,
	adErrDataOverflow = 3721,
	adErrSchemaViolation = 3722,
	adErrSignMismatch = 3723,
	adErrCantConvertvalue = 3724,
	adErrCantCreate = 3725,
	adErrColumnNotOnThisRow = 3726,
	adErrURLDoesNotExist = 3727,
	adErrTreePermissionDenied = 3728,
	adErrInvalidURL = 3729,
	adErrResourceLocked = 3730,
	adErrResourceExists = 3731,
	adErrCannotComplete = 3732,
	adErrVolumeNotFound = 3733,
	adErrOutOfSpace = 3734,
	adErrResourceOutOfScope = 3735,
	adErrUnavailable = 3736,
	adErrURLNamedRowDoesNotExist = 3737,
	adErrDelResOutOfScope = 3738,
	adErrPropInvalidColumn = 3739,
	adErrPropInvalidOption = 3740,
	adErrPropInvalidValue = 3741,
	adErrPropConflicting = 3742,
	adErrPropNotAllSettable = 3743,
	adErrPropNotSet = 3744,
	adErrPropNotSettable = 3745,
	adErrPropNotSupported = 3746,
	adErrCatalogNotSet = 3747,
	adErrCantChangeConnection = 3748,
	adErrFieldsUpdateFailed = 3749,
	adErrDenyNotSupported = 3750,
	adErrDenyTypeNotSupported = 3751,
	adErrProviderNotSpecified = 3753,
	adErrConnectionStringTooLong = 3754
};

enum ParameterAttributesEnum
{
	adParamSigned = 16,
	adParamNullable = 64,
	adParamLong = 128
};

enum ParameterDirectionEnum
{
	adParamUnknown = 0,
	adParamInput = 1,
	adParamOutput = 2,
	adParamInputOutput = 3,
	adParamReturnValue = 4
};

enum CommandTypeEnum
{
	adCmdUnspecified = -1,
	adCmdUnknown = 8,
	adCmdText = 1,
	adCmdTable = 2,
	adCmdStoredProc = 4,
	adCmdFile = 256,
	adCmdTableDirect = 512
};

enum EventStatusEnum
{
	adStatusOK = 1,
	adStatusErrorsOccurred = 2,
	adStatusCantDeny = 3,
	adStatusCancel = 4,
	adStatusUnwantedEvent = 5
};

enum EventReasonEnum
{
	adRsnAddNew = 1,
	adRsnDelete = 2,
	adRsnUpdate = 3,
	adRsnUndoUpdate = 4,
	adRsnUndoAddNew = 5,
	adRsnUndoDelete = 6,
	adRsnRequery = 7,
	adRsnResynch = 8,
	adRsnClose = 9,
	adRsnMove = 10,
	adRsnFirstChange = 11,
	adRsnMoveFirst = 12,
	adRsnMoveNext = 13,
	adRsnMovePrevious = 14,
	adRsnMoveLast = 15
};

enum SchemaEnum
{
	adSchemaProviderSpecific = -1,
	adSchemaAsserts = 0,
	adSchemaCatalogs = 1,
	adSchemaCharacterSets = 2,
	adSchemaCollations = 3,
	adSchemaColumns = 4,
	adSchemaCheckConstraints = 5,
	adSchemaConstraintColumnUsage = 6,
	adSchemaConstraintTableUsage = 7,
	adSchemaKeyColumnUsage = 8,
	adSchemaReferentialContraints = 9,
	adSchemaReferentialConstraints = 9,
	adSchemaTableConstraints = 10,
	adSchemaColumnsDomainUsage = 11,
	adSchemaIndexes = 12,
	adSchemaColumnPrivileges = 13,
	adSchemaTablePrivileges = 14,
	adSchemaUsagePrivileges = 15,
	adSchemaProcedures = 16,
	adSchemaSchemata = 17,
	adSchemaSQLLanguages = 18,
	adSchemaStatistics = 19,
	adSchemaTables = 20,
	adSchemaTranslations = 21,
	adSchemaProviderTypes = 22,
	adSchemaViews = 23,
	adSchemaViewColumnUsage = 24,
	adSchemaViewTableUsage = 25,
	adSchemaProcedureParameters = 26,
	adSchemaForeignKeys = 27,
	adSchemaPrimaryKeys = 28,
	adSchemaProcedureColumns = 29,
	adSchemaDBInfoKeywords = 30,
	adSchemaDBInfoLiterals = 31,
	adSchemaCubes = 32,
	adSchemaDimensions = 33,
	adSchemaHierarchies = 34,
	adSchemaLevels = 35,
	adSchemaMeasures = 36,
	adSchemaProperties = 37,
	adSchemaMembers = 38,
	adSchemaTrustees = 39,
	adSchemaFunctions = 40,
	adSchemaActions = 41,
	adSchemaCommands = 42,
	adSchemaSets = 43
};

enum FieldStatusEnum
{
	adFieldOK = 0,
	adFieldCantConvertValue = 2,
	adFieldIsNull = 3,
	adFieldTruncated = 4,
	adFieldSignMismatch = 5,
	adFieldDataOverflow = 6,
	adFieldCantCreate = 7,
	adFieldUnavailable = 8,
	adFieldPermissionDenied = 9,
	adFieldIntegrityViolation = 10,
	adFieldSchemaViolation = 11,
	adFieldBadStatus = 12,
	adFieldDefault = 13,
	adFieldIgnore = 15,
	adFieldDoesNotExist = 16,
	adFieldInvalidURL = 17,
	adFieldResourceLocked = 18,
	adFieldResourceExists = 19,
	adFieldCannotComplete = 20,
	adFieldVolumeNotFound = 21,
	adFieldOutOfSpace = 22,
	adFieldCannotDeleteSource = 23,
	adFieldReadOnly = 24,
	adFieldResourceOutOfScope = 25,
	adFieldAlreadyExists = 26,
	adFieldPendingInsert = 65536,
	adFieldPendingDelete = 131072,
	adFieldPendingChange = 262144,
	adFieldPendingUnknown = 524288,
	adFieldPendingUnknownDelete = 1048576
};

enum SeekEnum
{
	adSeekFirstEQ = 1,
	adSeekLastEQ = 2,
	adSeekAfterEQ = 4,
	adSeekAfter = 8,
	adSeekBeforeEQ = 16,
	adSeekBefore = 32
};

enum ADCPROP_UPDATECRITERIA_ENUM
{
	adCriteriaKey = 0,
	adCriteriaAllCols = 1,
	adCriteriaUpdCols = 2,
	adCriteriaTimeStamp = 3
};

enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
{
	adPriorityLowest = 1,
	adPriorityBelowNormal = 2,
	adPriorityNormal = 3,
	adPriorityAboveNormal = 4,
	adPriorityHighest = 5
};

enum ADCPROP_AUTORECALC_ENUM
{
	adRecalcUpFront = 0,
	adRecalcAlways = 1
};

enum ADCPROP_UPDATERESYNC_ENUM
{
	adResyncNone = 0,
	adResyncAutoIncrement = 1,
	adResyncConflicts = 2,
	adResyncUpdates = 4,
	adResyncInserts = 8,
	adResyncAll = 15
};

enum MoveRecordOptionsEnum
{
	adMoveUnspecified = -1,
	adMoveOverWrite = 1,
	adMoveDontUpdateLinks = 2,
	adMoveAllowEmulation = 4
};

enum CopyRecordOptionsEnum
{
	adCopyUnspecified = -1,
	adCopyOverWrite = 1,
	adCopyAllowEmulation = 4,
	adCopyNonRecursive = 2
};

enum StreamTypeEnum
{
	adTypeBinary = 1,
	adTypeText = 2
};

enum LineSeparatorEnum
{
	adLF = 10,
	adCR = 13,
	adCRLF = -1
};

enum StreamOpenOptionsEnum
{
	adOpenStreamUnspecified = -1,
	adOpenStreamAsync = 1,
	adOpenStreamFromRecord = 4
};

enum StreamWriteEnum
{
	adWriteChar = 0,
	adWriteLine = 1,
	stWriteChar = 0,
	stWriteLine = 1
};

enum SaveOptionsEnum
{
	adSaveCreateNotExist = 1,
	adSaveCreateOverWrite = 2
};

enum FieldEnum
{
	adDefaultStream = -1,
	adRecordURL = -2
};

enum StreamReadEnum
{
	adReadAll = -1,
	adReadLine = -2
};

enum RecordTypeEnum
{
	adSimpleRecord = 0,
	adCollectionRecord = 1,
	adStructDoc = 2
};

struct __declspec(uuid("00000512-0000-0010-8000-00aa006d2ea4"))
_Collection : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetCount))
	long Count;

	//
	// Wrapper methods for error-handling
	//

	long GetCount ();
	IUnknownPtr _NewEnum ();
	HRESULT Refresh ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Count (
		long * c) = 0;
	virtual HRESULT __stdcall raw__NewEnum (
		IUnknown * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Refresh () = 0;
};

struct __declspec(uuid("00000513-0000-0010-8000-00aa006d2ea4"))
_DynaCollection : public _Collection
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT Append (
		IDispatch * Object);
	HRESULT Delete (
		const _variant_t & Index);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Append (
		IDispatch * Object) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Index) = 0;
};

struct __declspec(uuid("00000534-0000-0010-8000-00aa006d2ea4"))
_ADO : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetProperties))
	PropertiesPtr Properties;

	//
	// Wrapper methods for error-handling
	//

	PropertiesPtr GetProperties ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Properties (
		struct Properties * * ppvObject) = 0;
};

struct __declspec(uuid("00000504-0000-0010-8000-00aa006d2ea4"))
Properties : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	PropertyPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	PropertyPtr GetItem (
		const _variant_t & Index);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Index,
		struct Property * * ppvObject) = 0;
};

struct __declspec(uuid("00000503-0000-0010-8000-00aa006d2ea4"))
Property : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetValue, put = PutValue))
	_variant_t Value;
	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetType))
	enum DataTypeEnum Type;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	long Attributes;

	//
	// Wrapper methods for error-handling
	//

	_variant_t GetValue ();
	void PutValue (
		const _variant_t & pval);
	_bstr_t GetName ();
	enum DataTypeEnum GetType ();
	long GetAttributes ();
	void PutAttributes (
		long plAttributes);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Value (
		VARIANT * pval) = 0;
	virtual HRESULT __stdcall put_Value (
		VARIANT pval) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnum * ptype) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * plAttributes) = 0;
	virtual HRESULT __stdcall put_Attributes (
		long plAttributes) = 0;
};

struct __declspec(uuid("00000500-0000-0010-8000-00aa006d2ea4"))
Error : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetDescription))
	_bstr_t Description;
	__declspec(property(get = GetNumber))
	long Number;
	__declspec(property(get = GetSource))
	_bstr_t Source;
	__declspec(property(get = GetHelpFile))
	_bstr_t HelpFile;
	__declspec(property(get = GetHelpContext))
	long HelpContext;
	__declspec(property(get = GetSQLState))
	_bstr_t SQLState;
	__declspec(property(get = GetNativeError))
	long NativeError;

	//
	// Wrapper methods for error-handling
	//

	long GetNumber ();
	_bstr_t GetSource ();
	_bstr_t GetDescription ();
	_bstr_t GetHelpFile ();
	long GetHelpContext ();
	_bstr_t GetSQLState ();
	long GetNativeError ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Number (
		long * pl) = 0;
	virtual HRESULT __stdcall get_Source (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_Description (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_HelpFile (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_HelpContext (
		long * pl) = 0;
	virtual HRESULT __stdcall get_SQLState (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_NativeError (
		long * pl) = 0;
};

struct __declspec(uuid("00000501-0000-0010-8000-00aa006d2ea4"))
Errors : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	ErrorPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	ErrorPtr GetItem (
		const _variant_t & Index);
	HRESULT Clear ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Index,
		struct Error * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Clear () = 0;
};

struct __declspec(uuid("00000508-0000-0010-8000-00aa006d2ea4"))
Command15 : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetParameters))
	ParametersPtr Parameters;
	__declspec(property(get = GetActiveConnection, put = PutRefActiveConnection))
	_ConnectionPtr ActiveConnection;
	__declspec(property(get = GetCommandText, put = PutCommandText))
	_bstr_t CommandText;
	__declspec(property(get = GetCommandTimeout, put = PutCommandTimeout))
	long CommandTimeout;
	__declspec(property(get = GetPrepared, put = PutPrepared))
	VARIANT_BOOL Prepared;
	__declspec(property(get = GetCommandType, put = PutCommandType))
	enum CommandTypeEnum CommandType;
	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;

	//
	// Wrapper methods for error-handling
	//

	_ConnectionPtr GetActiveConnection ();
	void PutRefActiveConnection (
		struct _Connection * ppvObject);
	void PutActiveConnection (
		const _variant_t & ppvObject);
	_bstr_t GetCommandText ();
	void PutCommandText (
		_bstr_t pbstr);
	long GetCommandTimeout ();
	void PutCommandTimeout (
		long pl);
	VARIANT_BOOL GetPrepared ();
	void PutPrepared (
		VARIANT_BOOL pfPrepared);
	_RecordsetPtr Execute (
		VARIANT * RecordsAffected,
		VARIANT * Parameters,
		long Options);
	_ParameterPtr CreateParameter (
		_bstr_t Name,
		enum DataTypeEnum Type,
		enum ParameterDirectionEnum Direction,
		ADO_LONGPTR Size,
		const _variant_t & Value = vtMissing);
	ParametersPtr GetParameters ();
	void PutCommandType (
		enum CommandTypeEnum plCmdType);
	enum CommandTypeEnum GetCommandType ();
	_bstr_t GetName ();
	void PutName (
		_bstr_t pbstrName);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_ActiveConnection (
		struct _Connection * * ppvObject) = 0;
	virtual HRESULT __stdcall putref_ActiveConnection (
		struct _Connection * ppvObject) = 0;
	virtual HRESULT __stdcall put_ActiveConnection (
		VARIANT ppvObject) = 0;
	virtual HRESULT __stdcall get_CommandText (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall put_CommandText (
		BSTR pbstr) = 0;
	virtual HRESULT __stdcall get_CommandTimeout (
		long * pl) = 0;
	virtual HRESULT __stdcall put_CommandTimeout (
		long pl) = 0;
	virtual HRESULT __stdcall get_Prepared (
		VARIANT_BOOL * pfPrepared) = 0;
	virtual HRESULT __stdcall put_Prepared (
		VARIANT_BOOL pfPrepared) = 0;
	virtual HRESULT __stdcall raw_Execute (
		VARIANT * RecordsAffected,
		VARIANT * Parameters,
		long Options,
		struct _Recordset * * ppiRs) = 0;
	virtual HRESULT __stdcall raw_CreateParameter (
		BSTR Name,
		enum DataTypeEnum Type,
		enum ParameterDirectionEnum Direction,
		ADO_LONGPTR Size,
		VARIANT Value,
		struct _Parameter * * ppiprm) = 0;
	virtual HRESULT __stdcall get_Parameters (
		struct Parameters * * ppvObject) = 0;
	virtual HRESULT __stdcall put_CommandType (
		enum CommandTypeEnum plCmdType) = 0;
	virtual HRESULT __stdcall get_CommandType (
		enum CommandTypeEnum * plCmdType) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pbstrName) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pbstrName) = 0;
};

struct __declspec(uuid("00000515-0000-0010-8000-00aa006d2ea4"))
Connection15 : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetConnectionString, put = PutConnectionString))
	_bstr_t ConnectionString;
	__declspec(property(get = GetCommandTimeout, put = PutCommandTimeout))
	long CommandTimeout;
	__declspec(property(get = GetConnectionTimeout, put = PutConnectionTimeout))
	long ConnectionTimeout;
	__declspec(property(get = GetVersion))
	_bstr_t Version;
	__declspec(property(get = GetErrors))
	ErrorsPtr Errors;
	__declspec(property(get = GetDefaultDatabase, put = PutDefaultDatabase))
	_bstr_t DefaultDatabase;
	__declspec(property(get = GetIsolationLevel, put = PutIsolationLevel))
	enum IsolationLevelEnum IsolationLevel;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	long Attributes;
	__declspec(property(get = GetCursorLocation, put = PutCursorLocation))
	enum CursorLocationEnum CursorLocation;
	__declspec(property(get = GetMode, put = PutMode))
	enum ConnectModeEnum Mode;
	__declspec(property(get = GetProvider, put = PutProvider))
	_bstr_t Provider;
	__declspec(property(get = GetState))
	long State;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetConnectionString ();
	void PutConnectionString (
		_bstr_t pbstr);
	long GetCommandTimeout ();
	void PutCommandTimeout (
		long plTimeout);
	long GetConnectionTimeout ();
	void PutConnectionTimeout (
		long plTimeout);
	_bstr_t GetVersion ();
	HRESULT Close ();
	_RecordsetPtr Execute (
		_bstr_t CommandText,
		VARIANT * RecordsAffected,
		long Options);
	long BeginTrans ();
	HRESULT CommitTrans ();
	HRESULT RollbackTrans ();
	HRESULT Open (
		_bstr_t ConnectionString,
		_bstr_t UserID,
		_bstr_t Password,
		long Options);
	ErrorsPtr GetErrors ();
	_bstr_t GetDefaultDatabase ();
	void PutDefaultDatabase (
		_bstr_t pbstr);
	enum IsolationLevelEnum GetIsolationLevel ();
	void PutIsolationLevel (
		enum IsolationLevelEnum Level);
	long GetAttributes ();
	void PutAttributes (
		long plAttr);
	enum CursorLocationEnum GetCursorLocation ();
	void PutCursorLocation (
		enum CursorLocationEnum plCursorLoc);
	enum ConnectModeEnum GetMode ();
	void PutMode (
		enum ConnectModeEnum plMode);
	_bstr_t GetProvider ();
	void PutProvider (
		_bstr_t pbstr);
	long GetState ();
	_RecordsetPtr OpenSchema (
		enum SchemaEnum Schema,
		const _variant_t & Restrictions = vtMissing,
		const _variant_t & SchemaID = vtMissing);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_ConnectionString (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall put_ConnectionString (
		BSTR pbstr) = 0;
	virtual HRESULT __stdcall get_CommandTimeout (
		long * plTimeout) = 0;
	virtual HRESULT __stdcall put_CommandTimeout (
		long plTimeout) = 0;
	virtual HRESULT __stdcall get_ConnectionTimeout (
		long * plTimeout) = 0;
	virtual HRESULT __stdcall put_ConnectionTimeout (
		long plTimeout) = 0;
	virtual HRESULT __stdcall get_Version (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall raw_Close () = 0;
	virtual HRESULT __stdcall raw_Execute (
		BSTR CommandText,
		VARIANT * RecordsAffected,
		long Options,
		struct _Recordset * * ppiRset) = 0;
	virtual HRESULT __stdcall raw_BeginTrans (
		long * TransactionLevel) = 0;
	virtual HRESULT __stdcall raw_CommitTrans () = 0;
	virtual HRESULT __stdcall raw_RollbackTrans () = 0;
	virtual HRESULT __stdcall raw_Open (
		BSTR ConnectionString,
		BSTR UserID,
		BSTR Password,
		long Options) = 0;
	virtual HRESULT __stdcall get_Errors (
		struct Errors * * ppvObject) = 0;
	virtual HRESULT __stdcall get_DefaultDatabase (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall put_DefaultDatabase (
		BSTR pbstr) = 0;
	virtual HRESULT __stdcall get_IsolationLevel (
		enum IsolationLevelEnum * Level) = 0;
	virtual HRESULT __stdcall put_IsolationLevel (
		enum IsolationLevelEnum Level) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * plAttr) = 0;
	virtual HRESULT __stdcall put_Attributes (
		long plAttr) = 0;
	virtual HRESULT __stdcall get_CursorLocation (
		enum CursorLocationEnum * plCursorLoc) = 0;
	virtual HRESULT __stdcall put_CursorLocation (
		enum CursorLocationEnum plCursorLoc) = 0;
	virtual HRESULT __stdcall get_Mode (
		enum ConnectModeEnum * plMode) = 0;
	virtual HRESULT __stdcall put_Mode (
		enum ConnectModeEnum plMode) = 0;
	virtual HRESULT __stdcall get_Provider (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall put_Provider (
		BSTR pbstr) = 0;
	virtual HRESULT __stdcall get_State (
		long * plObjState) = 0;
	virtual HRESULT __stdcall raw_OpenSchema (
		enum SchemaEnum Schema,
		VARIANT Restrictions,
		VARIANT SchemaID,
		struct _Recordset * * pprset) = 0;
};

struct __declspec(uuid("00000550-0000-0010-8000-00aa006d2ea4"))
_Connection : public Connection15
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT Cancel ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Cancel () = 0;
};

struct __declspec(uuid("0000050e-0000-0010-8000-00aa006d2ea4"))
Recordset15 : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetFields))
	FieldsPtr Fields;
	__declspec(property(get = GetPageSize, put = PutPageSize))
	long PageSize;
	__declspec(property(get = GetPageCount))
	ADO_LONGPTR PageCount;
	__declspec(property(get = GetCursorLocation, put = PutCursorLocation))
	enum CursorLocationEnum CursorLocation;
	__declspec(property(get = GetState))
	long State;
	__declspec(property(get = GetMarshalOptions, put = PutMarshalOptions))
	enum MarshalOptionsEnum MarshalOptions;
	__declspec(property(get = GetCollect, put = PutCollect))
	_variant_t Collect[];
	__declspec(property(get = GetEditMode))
	enum EditModeEnum EditMode;
	__declspec(property(get = GetStatus))
	long Status;
	__declspec(property(get = GetFilter, put = PutFilter))
	_variant_t Filter;
	__declspec(property(get = GetSort, put = PutSort))
	_bstr_t Sort;
	__declspec(property(get = GetAbsolutePosition, put = PutAbsolutePosition))
	PositionEnum_Param AbsolutePosition;
	__declspec(property(get = GetbBOF))
	VARIANT_BOOL bBOF;
	__declspec(property(get = GetBookmark, put = PutBookmark))
	_variant_t Bookmark;
	__declspec(property(get = GetCacheSize, put = PutCacheSize))
	long CacheSize;
	__declspec(property(get = GetCursorType, put = PutCursorType))
	enum CursorTypeEnum CursorType;
	__declspec(property(get = GetbEOF))
	VARIANT_BOOL bEOF;
	__declspec(property(get = GetAbsolutePage, put = PutAbsolutePage))
	PositionEnum_Param AbsolutePage;
	__declspec(property(get = GetLockType, put = PutLockType))
	enum LockTypeEnum LockType;
	__declspec(property(get = GetMaxRecords, put = PutMaxRecords))
	ADO_LONGPTR MaxRecords;
	__declspec(property(get = GetRecordCount))
	ADO_LONGPTR RecordCount;

	//
	// Wrapper methods for error-handling
	//

	PositionEnum_Param GetAbsolutePosition ();
	void PutAbsolutePosition (
		PositionEnum_Param pl);
	void PutRefActiveConnection (
		IDispatch * pvar);
	void PutActiveConnection (
		const _variant_t & pvar);
	_variant_t GetActiveConnection ();
	VARIANT_BOOL GetbBOF ();
	_variant_t GetBookmark ();
	void PutBookmark (
		const _variant_t & pvBookmark);
	long GetCacheSize ();
	void PutCacheSize (
		long pl);
	enum CursorTypeEnum GetCursorType ();
	void PutCursorType (
		enum CursorTypeEnum plCursorType);
	VARIANT_BOOL GetbEOF ();
	FieldsPtr GetFields ();
	enum LockTypeEnum GetLockType ();
	void PutLockType (
		enum LockTypeEnum plLockType);
	ADO_LONGPTR GetMaxRecords ();
	void PutMaxRecords (
		ADO_LONGPTR plMaxRecords);
	ADO_LONGPTR GetRecordCount ();
	void PutRefSource (
		IDispatch * pvSource);
	void PutSource (
		_bstr_t pvSource);
	_variant_t GetSource ();
	HRESULT AddNew (
		const _variant_t & FieldList = vtMissing,
		const _variant_t & Values = vtMissing);
	HRESULT CancelUpdate ();
	HRESULT Close ();
	HRESULT Delete (
		enum AffectEnum AffectRecords);
	_variant_t GetRows (
		long Rows,
		const _variant_t & Start = vtMissing,
		const _variant_t & Fields = vtMissing);
	HRESULT Move (
		ADO_LONGPTR NumRecords,
		const _variant_t & Start = vtMissing);
	HRESULT MoveNext ();
	HRESULT MovePrevious ();
	HRESULT MoveFirst ();
	HRESULT MoveLast ();
	HRESULT Open (
		const _variant_t & Source,
		const _variant_t & ActiveConnection,
		enum CursorTypeEnum CursorType,
		enum LockTypeEnum LockType,
		long Options);
	HRESULT Requery (
		long Options);
	HRESULT _xResync (
		enum AffectEnum AffectRecords);
	HRESULT Update (
		const _variant_t & Fields = vtMissing,
		const _variant_t & Values = vtMissing);
	PositionEnum_Param GetAbsolutePage ();
	void PutAbsolutePage (
		PositionEnum_Param pl);
	enum EditModeEnum GetEditMode ();
	_variant_t GetFilter ();
	void PutFilter (
		const _variant_t & Criteria);
	ADO_LONGPTR GetPageCount ();
	long GetPageSize ();
	void PutPageSize (
		long pl);
	_bstr_t GetSort ();
	void PutSort (
		_bstr_t Criteria);
	long GetStatus ();
	long GetState ();
	_RecordsetPtr _xClone ();
	HRESULT UpdateBatch (
		enum AffectEnum AffectRecords);
	HRESULT CancelBatch (
		enum AffectEnum AffectRecords);
	enum CursorLocationEnum GetCursorLocation ();
	void PutCursorLocation (
		enum CursorLocationEnum plCursorLoc);
	_RecordsetPtr NextRecordset (
		VARIANT * RecordsAffected);
	VARIANT_BOOL Supports (
		enum CursorOptionEnum CursorOptions);
	_variant_t GetCollect (
		const _variant_t & Index);
	void PutCollect (
		const _variant_t & Index,
		const _variant_t & pvar);
	enum MarshalOptionsEnum GetMarshalOptions ();
	void PutMarshalOptions (
		enum MarshalOptionsEnum peMarshal);
	HRESULT Find (
		_bstr_t Criteria,
		ADO_LONGPTR SkipRecords,
		enum SearchDirectionEnum SearchDirection,
		const _variant_t & Start = vtMissing);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_AbsolutePosition (
		PositionEnum_Param * pl) = 0;
	virtual HRESULT __stdcall put_AbsolutePosition (
		PositionEnum_Param pl) = 0;
	virtual HRESULT __stdcall putref_ActiveConnection (
		IDispatch * pvar) = 0;
	virtual HRESULT __stdcall put_ActiveConnection (
		VARIANT pvar) = 0;
	virtual HRESULT __stdcall get_ActiveConnection (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_bBOF (
		VARIANT_BOOL * pb) = 0;
	virtual HRESULT __stdcall get_Bookmark (
		VARIANT * pvBookmark) = 0;
	virtual HRESULT __stdcall put_Bookmark (
		VARIANT pvBookmark) = 0;
	virtual HRESULT __stdcall get_CacheSize (
		long * pl) = 0;
	virtual HRESULT __stdcall put_CacheSize (
		long pl) = 0;
	virtual HRESULT __stdcall get_CursorType (
		enum CursorTypeEnum * plCursorType) = 0;
	virtual HRESULT __stdcall put_CursorType (
		enum CursorTypeEnum plCursorType) = 0;
	virtual HRESULT __stdcall get_bEOF (
		VARIANT_BOOL * pb) = 0;
	virtual HRESULT __stdcall get_Fields (
		struct Fields * * ppvObject) = 0;
	virtual HRESULT __stdcall get_LockType (
		enum LockTypeEnum * plLockType) = 0;
	virtual HRESULT __stdcall put_LockType (
		enum LockTypeEnum plLockType) = 0;
	virtual HRESULT __stdcall get_MaxRecords (
		ADO_LONGPTR * plMaxRecords) = 0;
	virtual HRESULT __stdcall put_MaxRecords (
		ADO_LONGPTR plMaxRecords) = 0;
	virtual HRESULT __stdcall get_RecordCount (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall putref_Source (
		IDispatch * pvSource) = 0;
	virtual HRESULT __stdcall put_Source (
		BSTR pvSource) = 0;
	virtual HRESULT __stdcall get_Source (
		VARIANT * pvSource) = 0;
	virtual HRESULT __stdcall raw_AddNew (
		VARIANT FieldList = vtMissing,
		VARIANT Values = vtMissing) = 0;
	virtual HRESULT __stdcall raw_CancelUpdate () = 0;
	virtual HRESULT __stdcall raw_Close () = 0;
	virtual HRESULT __stdcall raw_Delete (
		enum AffectEnum AffectRecords) = 0;
	virtual HRESULT __stdcall raw_GetRows (
		long Rows,
		VARIANT Start,
		VARIANT Fields,
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall raw_Move (
		ADO_LONGPTR NumRecords,
		VARIANT Start = vtMissing) = 0;
	virtual HRESULT __stdcall raw_MoveNext () = 0;
	virtual HRESULT __stdcall raw_MovePrevious () = 0;
	virtual HRESULT __stdcall raw_MoveFirst () = 0;
	virtual HRESULT __stdcall raw_MoveLast () = 0;
	virtual HRESULT __stdcall raw_Open (
		VARIANT Source,
		VARIANT ActiveConnection,
		enum CursorTypeEnum CursorType,
		enum LockTypeEnum LockType,
		long Options) = 0;
	virtual HRESULT __stdcall raw_Requery (
		long Options) = 0;
	virtual HRESULT __stdcall raw__xResync (
		enum AffectEnum AffectRecords) = 0;
	virtual HRESULT __stdcall raw_Update (
		VARIANT Fields = vtMissing,
		VARIANT Values = vtMissing) = 0;
	virtual HRESULT __stdcall get_AbsolutePage (
		PositionEnum_Param * pl) = 0;
	virtual HRESULT __stdcall put_AbsolutePage (
		PositionEnum_Param pl) = 0;
	virtual HRESULT __stdcall get_EditMode (
		enum EditModeEnum * pl) = 0;
	virtual HRESULT __stdcall get_Filter (
		VARIANT * Criteria) = 0;
	virtual HRESULT __stdcall put_Filter (
		VARIANT Criteria) = 0;
	virtual HRESULT __stdcall get_PageCount (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall get_PageSize (
		long * pl) = 0;
	virtual HRESULT __stdcall put_PageSize (
		long pl) = 0;
	virtual HRESULT __stdcall get_Sort (
		BSTR * Criteria) = 0;
	virtual HRESULT __stdcall put_Sort (
		BSTR Criteria) = 0;
	virtual HRESULT __stdcall get_Status (
		long * pl) = 0;
	virtual HRESULT __stdcall get_State (
		long * plObjState) = 0;
	virtual HRESULT __stdcall raw__xClone (
		struct _Recordset * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_UpdateBatch (
		enum AffectEnum AffectRecords) = 0;
	virtual HRESULT __stdcall raw_CancelBatch (
		enum AffectEnum AffectRecords) = 0;
	virtual HRESULT __stdcall get_CursorLocation (
		enum CursorLocationEnum * plCursorLoc) = 0;
	virtual HRESULT __stdcall put_CursorLocation (
		enum CursorLocationEnum plCursorLoc) = 0;
	virtual HRESULT __stdcall raw_NextRecordset (
		VARIANT * RecordsAffected,
		struct _Recordset * * ppiRs) = 0;
	virtual HRESULT __stdcall raw_Supports (
		enum CursorOptionEnum CursorOptions,
		VARIANT_BOOL * pb) = 0;
	virtual HRESULT __stdcall get_Collect (
		VARIANT Index,
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_Collect (
		VARIANT Index,
		VARIANT pvar) = 0;
	virtual HRESULT __stdcall get_MarshalOptions (
		enum MarshalOptionsEnum * peMarshal) = 0;
	virtual HRESULT __stdcall put_MarshalOptions (
		enum MarshalOptionsEnum peMarshal) = 0;
	virtual HRESULT __stdcall raw_Find (
		BSTR Criteria,
		ADO_LONGPTR SkipRecords,
		enum SearchDirectionEnum SearchDirection,
		VARIANT Start = vtMissing) = 0;
};

struct __declspec(uuid("0000054f-0000-0010-8000-00aa006d2ea4"))
Recordset20 : public Recordset15
{
	//
	// Property data
	//

	__declspec(property(get = GetDataSource, put = PutRefDataSource))
	IUnknownPtr DataSource;
	__declspec(property(get = GetActiveCommand))
	IDispatchPtr ActiveCommand;
	__declspec(property(get = GetStayInSync, put = PutStayInSync))
	VARIANT_BOOL StayInSync;
	__declspec(property(get = GetDataMember, put = PutDataMember))
	_bstr_t DataMember;

	//
	// Wrapper methods for error-handling
	//

	HRESULT Cancel ();
	IUnknownPtr GetDataSource ();
	void PutRefDataSource (
		IUnknown * ppunkDataSource);
	HRESULT _xSave (
		_bstr_t FileName,
		enum PersistFormatEnum PersistFormat);
	IDispatchPtr GetActiveCommand ();
	void PutStayInSync (
		VARIANT_BOOL pbStayInSync);
	VARIANT_BOOL GetStayInSync ();
	_bstr_t GetString (
		enum StringFormatEnum StringFormat,
		long NumRows,
		_bstr_t ColumnDelimeter,
		_bstr_t RowDelimeter,
		_bstr_t NullExpr);
	_bstr_t GetDataMember ();
	void PutDataMember (
		_bstr_t pbstrDataMember);
	enum CompareEnum CompareBookmarks (
		const _variant_t & Bookmark1,
		const _variant_t & Bookmark2);
	_RecordsetPtr Clone (
		enum LockTypeEnum LockType);
	HRESULT Resync (
		enum AffectEnum AffectRecords,
		enum ResyncEnum ResyncValues);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Cancel () = 0;
	virtual HRESULT __stdcall get_DataSource (
		IUnknown * * ppunkDataSource) = 0;
	virtual HRESULT __stdcall putref_DataSource (
		IUnknown * ppunkDataSource) = 0;
	virtual HRESULT __stdcall raw__xSave (
		BSTR FileName,
		enum PersistFormatEnum PersistFormat) = 0;
	virtual HRESULT __stdcall get_ActiveCommand (
		IDispatch * * ppCmd) = 0;
	virtual HRESULT __stdcall put_StayInSync (
		VARIANT_BOOL pbStayInSync) = 0;
	virtual HRESULT __stdcall get_StayInSync (
		VARIANT_BOOL * pbStayInSync) = 0;
	virtual HRESULT __stdcall raw_GetString (
		enum StringFormatEnum StringFormat,
		long NumRows,
		BSTR ColumnDelimeter,
		BSTR RowDelimeter,
		BSTR NullExpr,
		BSTR * pRetString) = 0;
	virtual HRESULT __stdcall get_DataMember (
		BSTR * pbstrDataMember) = 0;
	virtual HRESULT __stdcall put_DataMember (
		BSTR pbstrDataMember) = 0;
	virtual HRESULT __stdcall raw_CompareBookmarks (
		VARIANT Bookmark1,
		VARIANT Bookmark2,
		enum CompareEnum * pCompare) = 0;
	virtual HRESULT __stdcall raw_Clone (
		enum LockTypeEnum LockType,
		struct _Recordset * * ppvObject) = 0;
	virtual HRESULT __stdcall raw_Resync (
		enum AffectEnum AffectRecords,
		enum ResyncEnum ResyncValues) = 0;
};

struct __declspec(uuid("00000555-0000-0010-8000-00aa006d2ea4"))
Recordset21 : public Recordset20
{
	//
	// Property data
	//

	__declspec(property(get = GetIndex, put = PutIndex))
	_bstr_t Index;

	//
	// Wrapper methods for error-handling
	//

	HRESULT Seek (
		const _variant_t & KeyValues,
		enum SeekEnum SeekOption);
	void PutIndex (
		_bstr_t pbstrIndex);
	_bstr_t GetIndex ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Seek (
		VARIANT KeyValues,
		enum SeekEnum SeekOption) = 0;
	virtual HRESULT __stdcall put_Index (
		BSTR pbstrIndex) = 0;
	virtual HRESULT __stdcall get_Index (
		BSTR * pbstrIndex) = 0;
};

struct __declspec(uuid("00000556-0000-0010-8000-00aa006d2ea4"))
_Recordset : public Recordset21
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT Save (
		const _variant_t & Destination,
		enum PersistFormatEnum PersistFormat);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Save (
		VARIANT Destination,
		enum PersistFormatEnum PersistFormat) = 0;
};

struct __declspec(uuid("00000506-0000-0010-8000-00aa006d2ea4"))
Fields15 : public _Collection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	FieldPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	FieldPtr GetItem (
		const _variant_t & Index);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Index,
		struct Field * * ppvObject) = 0;
};

struct __declspec(uuid("0000054d-0000-0010-8000-00aa006d2ea4"))
Fields20 : public Fields15
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT _Append (
		_bstr_t Name,
		enum DataTypeEnum Type,
		ADO_LONGPTR DefinedSize,
		enum FieldAttributeEnum Attrib);
	HRESULT Delete (
		const _variant_t & Index);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw__Append (
		BSTR Name,
		enum DataTypeEnum Type,
		ADO_LONGPTR DefinedSize,
		enum FieldAttributeEnum Attrib) = 0;
	virtual HRESULT __stdcall raw_Delete (
		VARIANT Index) = 0;
};

struct __declspec(uuid("00000564-0000-0010-8000-00aa006d2ea4"))
Fields : public Fields20
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT Append (
		_bstr_t Name,
		enum DataTypeEnum Type,
		ADO_LONGPTR DefinedSize,
		enum FieldAttributeEnum Attrib,
		const _variant_t & FieldValue = vtMissing);
	HRESULT Update ();
	HRESULT Resync (
		enum ResyncEnum ResyncValues);
	HRESULT CancelUpdate ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_Append (
		BSTR Name,
		enum DataTypeEnum Type,
		ADO_LONGPTR DefinedSize,
		enum FieldAttributeEnum Attrib,
		VARIANT FieldValue = vtMissing) = 0;
	virtual HRESULT __stdcall raw_Update () = 0;
	virtual HRESULT __stdcall raw_Resync (
		enum ResyncEnum ResyncValues) = 0;
	virtual HRESULT __stdcall raw_CancelUpdate () = 0;
};

struct __declspec(uuid("0000054c-0000-0010-8000-00aa006d2ea4"))
Field20 : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetValue, put = PutValue))
	_variant_t Value;
	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetType, put = PutType))
	enum DataTypeEnum Type;
	__declspec(property(get = GetDefinedSize, put = PutDefinedSize))
	ADO_LONGPTR DefinedSize;
	__declspec(property(get = GetOriginalValue))
	_variant_t OriginalValue;
	__declspec(property(get = GetUnderlyingValue))
	_variant_t UnderlyingValue;
	__declspec(property(get = GetActualSize))
	ADO_LONGPTR ActualSize;
	__declspec(property(get = GetPrecision, put = PutPrecision))
	unsigned char Precision;
	__declspec(property(get = GetNumericScale, put = PutNumericScale))
	unsigned char NumericScale;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	long Attributes;
	__declspec(property(get = GetDataFormat, put = PutRefDataFormat))
	IUnknownPtr DataFormat;

	//
	// Wrapper methods for error-handling
	//

	ADO_LONGPTR GetActualSize ();
	long GetAttributes ();
	ADO_LONGPTR GetDefinedSize ();
	_bstr_t GetName ();
	enum DataTypeEnum GetType ();
	_variant_t GetValue ();
	void PutValue (
		const _variant_t & pvar);
	unsigned char GetPrecision ();
	unsigned char GetNumericScale ();
	HRESULT AppendChunk (
		const _variant_t & Data);
	_variant_t GetChunk (
		long Length);
	_variant_t GetOriginalValue ();
	_variant_t GetUnderlyingValue ();
	IUnknownPtr GetDataFormat ();
	void PutRefDataFormat (
		IUnknown * ppiDF);
	void PutPrecision (
		unsigned char pbPrecision);
	void PutNumericScale (
		unsigned char pbNumericScale);
	void PutType (
		enum DataTypeEnum pDataType);
	void PutDefinedSize (
		ADO_LONGPTR pl);
	void PutAttributes (
		long pl);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_ActualSize (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * pl) = 0;
	virtual HRESULT __stdcall get_DefinedSize (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnum * pDataType) = 0;
	virtual HRESULT __stdcall get_Value (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_Value (
		VARIANT pvar) = 0;
	virtual HRESULT __stdcall get_Precision (
		unsigned char * pbPrecision) = 0;
	virtual HRESULT __stdcall get_NumericScale (
		unsigned char * pbNumericScale) = 0;
	virtual HRESULT __stdcall raw_AppendChunk (
		VARIANT Data) = 0;
	virtual HRESULT __stdcall raw_GetChunk (
		long Length,
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_OriginalValue (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_UnderlyingValue (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_DataFormat (
		IUnknown * * ppiDF) = 0;
	virtual HRESULT __stdcall putref_DataFormat (
		IUnknown * ppiDF) = 0;
	virtual HRESULT __stdcall put_Precision (
		unsigned char pbPrecision) = 0;
	virtual HRESULT __stdcall put_NumericScale (
		unsigned char pbNumericScale) = 0;
	virtual HRESULT __stdcall put_Type (
		enum DataTypeEnum pDataType) = 0;
	virtual HRESULT __stdcall put_DefinedSize (
		ADO_LONGPTR pl) = 0;
	virtual HRESULT __stdcall put_Attributes (
		long pl) = 0;
};

struct __declspec(uuid("00000569-0000-0010-8000-00aa006d2ea4"))
Field : public Field20
{
	//
	// Property data
	//

	__declspec(property(get = GetStatus))
	long Status;

	//
	// Wrapper methods for error-handling
	//

	long GetStatus ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Status (
		long * pFStatus) = 0;
};

struct __declspec(uuid("0000050c-0000-0010-8000-00aa006d2ea4"))
_Parameter : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetValue, put = PutValue))
	_variant_t Value;
	__declspec(property(get = GetName, put = PutName))
	_bstr_t Name;
	__declspec(property(get = GetType, put = PutType))
	enum DataTypeEnum Type;
	__declspec(property(get = GetDirection, put = PutDirection))
	enum ParameterDirectionEnum Direction;
	__declspec(property(get = GetPrecision, put = PutPrecision))
	unsigned char Precision;
	__declspec(property(get = GetNumericScale, put = PutNumericScale))
	unsigned char NumericScale;
	__declspec(property(get = GetSize, put = PutSize))
	ADO_LONGPTR Size;
	__declspec(property(get = GetAttributes, put = PutAttributes))
	long Attributes;

	//
	// Wrapper methods for error-handling
	//

	_bstr_t GetName ();
	void PutName (
		_bstr_t pbstr);
	_variant_t GetValue ();
	void PutValue (
		const _variant_t & pvar);
	enum DataTypeEnum GetType ();
	void PutType (
		enum DataTypeEnum psDataType);
	void PutDirection (
		enum ParameterDirectionEnum plParmDirection);
	enum ParameterDirectionEnum GetDirection ();
	void PutPrecision (
		unsigned char pbPrecision);
	unsigned char GetPrecision ();
	void PutNumericScale (
		unsigned char pbScale);
	unsigned char GetNumericScale ();
	void PutSize (
		ADO_LONGPTR pl);
	ADO_LONGPTR GetSize ();
	HRESULT AppendChunk (
		const _variant_t & Val);
	long GetAttributes ();
	void PutAttributes (
		long plParmAttribs);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Name (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall put_Name (
		BSTR pbstr) = 0;
	virtual HRESULT __stdcall get_Value (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_Value (
		VARIANT pvar) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnum * psDataType) = 0;
	virtual HRESULT __stdcall put_Type (
		enum DataTypeEnum psDataType) = 0;
	virtual HRESULT __stdcall put_Direction (
		enum ParameterDirectionEnum plParmDirection) = 0;
	virtual HRESULT __stdcall get_Direction (
		enum ParameterDirectionEnum * plParmDirection) = 0;
	virtual HRESULT __stdcall put_Precision (
		unsigned char pbPrecision) = 0;
	virtual HRESULT __stdcall get_Precision (
		unsigned char * pbPrecision) = 0;
	virtual HRESULT __stdcall put_NumericScale (
		unsigned char pbScale) = 0;
	virtual HRESULT __stdcall get_NumericScale (
		unsigned char * pbScale) = 0;
	virtual HRESULT __stdcall put_Size (
		ADO_LONGPTR pl) = 0;
	virtual HRESULT __stdcall get_Size (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall raw_AppendChunk (
		VARIANT Val) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * plParmAttribs) = 0;
	virtual HRESULT __stdcall put_Attributes (
		long plParmAttribs) = 0;
};

struct __declspec(uuid("0000050d-0000-0010-8000-00aa006d2ea4"))
Parameters : public _DynaCollection
{
	//
	// Property data
	//

	__declspec(property(get = GetItem))
	_ParameterPtr Item[];

	//
	// Wrapper methods for error-handling
	//

	_ParameterPtr GetItem (
		const _variant_t & Index);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Item (
		VARIANT Index,
		struct _Parameter * * ppvObject) = 0;
};

struct __declspec(uuid("0000054e-0000-0010-8000-00aa006d2ea4"))
Command25 : public Command15
{
	//
	// Property data
	//

	__declspec(property(get = GetState))
	long State;

	//
	// Wrapper methods for error-handling
	//

	long GetState ();
	HRESULT Cancel ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_State (
		long * plObjState) = 0;
	virtual HRESULT __stdcall raw_Cancel () = 0;
};

struct __declspec(uuid("b08400bd-f9d1-4d02-b856-71d5dba123e9"))
_Command : public Command25
{
	//
	// Property data
	//

	__declspec(property(get = GetDialect, put = PutDialect))
	_bstr_t Dialect;
	__declspec(property(get = GetNamedParameters, put = PutNamedParameters))
	VARIANT_BOOL NamedParameters;

	//
	// Wrapper methods for error-handling
	//

	void PutRefCommandStream (
		IUnknown * pvStream);
	_variant_t GetCommandStream ();
	void PutDialect (
		_bstr_t pbstrDialect);
	_bstr_t GetDialect ();
	void PutNamedParameters (
		VARIANT_BOOL pfNamedParameters);
	VARIANT_BOOL GetNamedParameters ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall putref_CommandStream (
		IUnknown * pvStream) = 0;
	virtual HRESULT __stdcall get_CommandStream (
		VARIANT * pvStream) = 0;
	virtual HRESULT __stdcall put_Dialect (
		BSTR pbstrDialect) = 0;
	virtual HRESULT __stdcall get_Dialect (
		BSTR * pbstrDialect) = 0;
	virtual HRESULT __stdcall put_NamedParameters (
		VARIANT_BOOL pfNamedParameters) = 0;
	virtual HRESULT __stdcall get_NamedParameters (
		VARIANT_BOOL * pfNamedParameters) = 0;
};

struct __declspec(uuid("00000402-0000-0010-8000-00aa006d2ea4"))
ConnectionEventsVt : public IUnknown
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT InfoMessage (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT BeginTransComplete (
		long TransactionLevel,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT CommitTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT RollbackTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT WillExecute (
		BSTR * Source,
		enum CursorTypeEnum * CursorType,
		enum LockTypeEnum * LockType,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection);
	HRESULT ExecuteComplete (
		long RecordsAffected,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection);
	HRESULT WillConnect (
		BSTR * ConnectionString,
		BSTR * UserID,
		BSTR * Password,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT ConnectComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT Disconnect (
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_InfoMessage (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_BeginTransComplete (
		long TransactionLevel,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_CommitTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_RollbackTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_WillExecute (
		BSTR * Source,
		enum CursorTypeEnum * CursorType,
		enum LockTypeEnum * LockType,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_ExecuteComplete (
		long RecordsAffected,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_WillConnect (
		BSTR * ConnectionString,
		BSTR * UserID,
		BSTR * Password,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_ConnectComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
	virtual HRESULT __stdcall raw_Disconnect (
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection) = 0;
};

struct __declspec(uuid("00000403-0000-0010-8000-00aa006d2ea4"))
RecordsetEventsVt : public IUnknown
{
	//
	// Wrapper methods for error-handling
	//

	HRESULT WillChangeField (
		long cFields,
		const _variant_t & Fields,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FieldChangeComplete (
		long cFields,
		const _variant_t & Fields,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillChangeRecord (
		enum EventReasonEnum adReason,
		long cRecords,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT RecordChangeComplete (
		enum EventReasonEnum adReason,
		long cRecords,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillChangeRecordset (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT RecordsetChangeComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillMove (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT MoveComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT EndOfRecordset (
		VARIANT_BOOL * fMoreData,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FetchProgress (
		long Progress,
		long MaxProgress,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FetchComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall raw_WillChangeField (
		long cFields,
		VARIANT Fields,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_FieldChangeComplete (
		long cFields,
		VARIANT Fields,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_WillChangeRecord (
		enum EventReasonEnum adReason,
		long cRecords,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_RecordChangeComplete (
		enum EventReasonEnum adReason,
		long cRecords,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_WillChangeRecordset (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_RecordsetChangeComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_WillMove (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_MoveComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_EndOfRecordset (
		VARIANT_BOOL * fMoreData,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_FetchProgress (
		long Progress,
		long MaxProgress,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
	virtual HRESULT __stdcall raw_FetchComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset) = 0;
};

struct __declspec(uuid("00000400-0000-0010-8000-00aa006d2ea4"))
ConnectionEvents : public IDispatch
{
	//
	// Wrapper methods for error-handling
	//

	// Methods:
	HRESULT InfoMessage (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT BeginTransComplete (
		long TransactionLevel,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT CommitTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT RollbackTransComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT WillExecute (
		BSTR * Source,
		enum CursorTypeEnum * CursorType,
		enum LockTypeEnum * LockType,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection);
	HRESULT ExecuteComplete (
		long RecordsAffected,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Command * pCommand,
		struct _Recordset * pRecordset,
		struct _Connection * pConnection);
	HRESULT WillConnect (
		BSTR * ConnectionString,
		BSTR * UserID,
		BSTR * Password,
		long * Options,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT ConnectComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
	HRESULT Disconnect (
		enum EventStatusEnum * adStatus,
		struct _Connection * pConnection);
};

struct __declspec(uuid("00000266-0000-0010-8000-00aa006d2ea4"))
RecordsetEvents : public IDispatch
{
	//
	// Wrapper methods for error-handling
	//

	// Methods:
	HRESULT WillChangeField (
		long cFields,
		const _variant_t & Fields,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FieldChangeComplete (
		long cFields,
		const _variant_t & Fields,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillChangeRecord (
		enum EventReasonEnum adReason,
		long cRecords,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT RecordChangeComplete (
		enum EventReasonEnum adReason,
		long cRecords,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillChangeRecordset (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT RecordsetChangeComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT WillMove (
		enum EventReasonEnum adReason,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT MoveComplete (
		enum EventReasonEnum adReason,
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT EndOfRecordset (
		VARIANT_BOOL * fMoreData,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FetchProgress (
		long Progress,
		long MaxProgress,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
	HRESULT FetchComplete (
		struct Error * pError,
		enum EventStatusEnum * adStatus,
		struct _Recordset * pRecordset);
};

struct __declspec(uuid("00000516-0000-0010-8000-00aa006d2ea4"))
ADOConnectionConstruction15 : public IUnknown
{
	//
	// Property data
	//

	__declspec(property(get = GetDSO))
	IUnknownPtr DSO;
	__declspec(property(get = GetSession))
	IUnknownPtr Session;

	//
	// Wrapper methods for error-handling
	//

	IUnknownPtr GetDSO ();
	IUnknownPtr GetSession ();
	HRESULT WrapDSOandSession (
		IUnknown * pDSO,
		IUnknown * pSession);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_DSO (
		IUnknown * * ppDSO) = 0;
	virtual HRESULT __stdcall get_Session (
		IUnknown * * ppSession) = 0;
	virtual HRESULT __stdcall raw_WrapDSOandSession (
		IUnknown * pDSO,
		IUnknown * pSession) = 0;
};

struct __declspec(uuid("00000551-0000-0010-8000-00aa006d2ea4"))
ADOConnectionConstruction : public ADOConnectionConstruction15
{};

struct __declspec(uuid("00000514-0000-0010-8000-00aa006d2ea4"))
Connection;
	// [ default ] interface _Connection
	// [ default, source ] dispinterface ConnectionEvents

struct __declspec(uuid("00000562-0000-0010-8000-00aa006d2ea4"))
_Record : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetFields))
	FieldsPtr Fields;
	__declspec(property(get = GetState))
	enum ObjectStateEnum State;
	__declspec(property(get = GetMode, put = PutMode))
	enum ConnectModeEnum Mode;
	__declspec(property(get = GetParentURL))
	_bstr_t ParentURL;
	__declspec(property(get = GetRecordType))
	enum RecordTypeEnum RecordType;

	//
	// Wrapper methods for error-handling
	//

	_variant_t GetActiveConnection ();
	void PutActiveConnection (
		_bstr_t pvar);
	void PutRefActiveConnection (
		struct _Connection * pvar);
	enum ObjectStateEnum GetState ();
	_variant_t GetSource ();
	void PutSource (
		_bstr_t pvar);
	void PutRefSource (
		IDispatch * pvar);
	enum ConnectModeEnum GetMode ();
	void PutMode (
		enum ConnectModeEnum pMode);
	_bstr_t GetParentURL ();
	_bstr_t MoveRecord (
		_bstr_t Source,
		_bstr_t Destination,
		_bstr_t UserName,
		_bstr_t Password,
		enum MoveRecordOptionsEnum Options,
		VARIANT_BOOL Async);
	_bstr_t CopyRecord (
		_bstr_t Source,
		_bstr_t Destination,
		_bstr_t UserName,
		_bstr_t Password,
		enum CopyRecordOptionsEnum Options,
		VARIANT_BOOL Async);
	HRESULT DeleteRecord (
		_bstr_t Source,
		VARIANT_BOOL Async);
	HRESULT Open (
		const _variant_t & Source,
		const _variant_t & ActiveConnection,
		enum ConnectModeEnum Mode,
		enum RecordCreateOptionsEnum CreateOptions,
		enum RecordOpenOptionsEnum Options,
		_bstr_t UserName,
		_bstr_t Password);
	HRESULT Close ();
	FieldsPtr GetFields ();
	enum RecordTypeEnum GetRecordType ();
	_RecordsetPtr GetChildren ();
	HRESULT Cancel ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_ActiveConnection (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_ActiveConnection (
		BSTR pvar) = 0;
	virtual HRESULT __stdcall putref_ActiveConnection (
		struct _Connection * pvar) = 0;
	virtual HRESULT __stdcall get_State (
		enum ObjectStateEnum * pState) = 0;
	virtual HRESULT __stdcall get_Source (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_Source (
		BSTR pvar) = 0;
	virtual HRESULT __stdcall putref_Source (
		IDispatch * pvar) = 0;
	virtual HRESULT __stdcall get_Mode (
		enum ConnectModeEnum * pMode) = 0;
	virtual HRESULT __stdcall put_Mode (
		enum ConnectModeEnum pMode) = 0;
	virtual HRESULT __stdcall get_ParentURL (
		BSTR * pbstrParentURL) = 0;
	virtual HRESULT __stdcall raw_MoveRecord (
		BSTR Source,
		BSTR Destination,
		BSTR UserName,
		BSTR Password,
		enum MoveRecordOptionsEnum Options,
		VARIANT_BOOL Async,
		BSTR * pbstrNewURL) = 0;
	virtual HRESULT __stdcall raw_CopyRecord (
		BSTR Source,
		BSTR Destination,
		BSTR UserName,
		BSTR Password,
		enum CopyRecordOptionsEnum Options,
		VARIANT_BOOL Async,
		BSTR * pbstrNewURL) = 0;
	virtual HRESULT __stdcall raw_DeleteRecord (
		BSTR Source,
		VARIANT_BOOL Async) = 0;
	virtual HRESULT __stdcall raw_Open (
		VARIANT Source,
		VARIANT ActiveConnection,
		enum ConnectModeEnum Mode,
		enum RecordCreateOptionsEnum CreateOptions,
		enum RecordOpenOptionsEnum Options,
		BSTR UserName,
		BSTR Password) = 0;
	virtual HRESULT __stdcall raw_Close () = 0;
	virtual HRESULT __stdcall get_Fields (
		struct Fields * * ppFlds) = 0;
	virtual HRESULT __stdcall get_RecordType (
		enum RecordTypeEnum * ptype) = 0;
	virtual HRESULT __stdcall raw_GetChildren (
		struct _Recordset * * pprset) = 0;
	virtual HRESULT __stdcall raw_Cancel () = 0;
};

struct __declspec(uuid("00000560-0000-0010-8000-00aa006d2ea4"))
Record;
	// [ default ] interface _Record

struct __declspec(uuid("00000565-0000-0010-8000-00aa006d2ea4"))
_Stream : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetSize))
	ADO_LONGPTR Size;
	__declspec(property(get = GetEOS))
	VARIANT_BOOL EOS;
	__declspec(property(get = GetPosition, put = PutPosition))
	ADO_LONGPTR Position;
	__declspec(property(get = GetType, put = PutType))
	enum StreamTypeEnum Type;
	__declspec(property(get = GetLineSeparator, put = PutLineSeparator))
	enum LineSeparatorEnum LineSeparator;
	__declspec(property(get = GetState))
	enum ObjectStateEnum State;
	__declspec(property(get = GetMode, put = PutMode))
	enum ConnectModeEnum Mode;
	__declspec(property(get = GetCharset, put = PutCharset))
	_bstr_t Charset;

	//
	// Wrapper methods for error-handling
	//

	ADO_LONGPTR GetSize ();
	VARIANT_BOOL GetEOS ();
	ADO_LONGPTR GetPosition ();
	void PutPosition (
		ADO_LONGPTR pPos);
	enum StreamTypeEnum GetType ();
	void PutType (
		enum StreamTypeEnum ptype);
	enum LineSeparatorEnum GetLineSeparator ();
	void PutLineSeparator (
		enum LineSeparatorEnum pLS);
	enum ObjectStateEnum GetState ();
	enum ConnectModeEnum GetMode ();
	void PutMode (
		enum ConnectModeEnum pMode);
	_bstr_t GetCharset ();
	void PutCharset (
		_bstr_t pbstrCharset);
	_variant_t Read (
		long NumBytes);
	HRESULT Open (
		const _variant_t & Source,
		enum ConnectModeEnum Mode,
		enum StreamOpenOptionsEnum Options,
		_bstr_t UserName,
		_bstr_t Password);
	HRESULT Close ();
	HRESULT SkipLine ();
	HRESULT Write (
		const _variant_t & Buffer);
	HRESULT SetEOS ();
	HRESULT CopyTo (
		struct _Stream * DestStream,
		ADO_LONGPTR CharNumber);
	HRESULT Flush ();
	HRESULT SaveToFile (
		_bstr_t FileName,
		enum SaveOptionsEnum Options);
	HRESULT LoadFromFile (
		_bstr_t FileName);
	_bstr_t ReadText (
		long NumChars);
	HRESULT WriteText (
		_bstr_t Data,
		enum StreamWriteEnum Options);
	HRESULT Cancel ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Size (
		ADO_LONGPTR * pSize) = 0;
	virtual HRESULT __stdcall get_EOS (
		VARIANT_BOOL * pEOS) = 0;
	virtual HRESULT __stdcall get_Position (
		ADO_LONGPTR * pPos) = 0;
	virtual HRESULT __stdcall put_Position (
		ADO_LONGPTR pPos) = 0;
	virtual HRESULT __stdcall get_Type (
		enum StreamTypeEnum * ptype) = 0;
	virtual HRESULT __stdcall put_Type (
		enum StreamTypeEnum ptype) = 0;
	virtual HRESULT __stdcall get_LineSeparator (
		enum LineSeparatorEnum * pLS) = 0;
	virtual HRESULT __stdcall put_LineSeparator (
		enum LineSeparatorEnum pLS) = 0;
	virtual HRESULT __stdcall get_State (
		enum ObjectStateEnum * pState) = 0;
	virtual HRESULT __stdcall get_Mode (
		enum ConnectModeEnum * pMode) = 0;
	virtual HRESULT __stdcall put_Mode (
		enum ConnectModeEnum pMode) = 0;
	virtual HRESULT __stdcall get_Charset (
		BSTR * pbstrCharset) = 0;
	virtual HRESULT __stdcall put_Charset (
		BSTR pbstrCharset) = 0;
	virtual HRESULT __stdcall raw_Read (
		long NumBytes,
		VARIANT * pval) = 0;
	virtual HRESULT __stdcall raw_Open (
		VARIANT Source,
		enum ConnectModeEnum Mode,
		enum StreamOpenOptionsEnum Options,
		BSTR UserName,
		BSTR Password) = 0;
	virtual HRESULT __stdcall raw_Close () = 0;
	virtual HRESULT __stdcall raw_SkipLine () = 0;
	virtual HRESULT __stdcall raw_Write (
		VARIANT Buffer) = 0;
	virtual HRESULT __stdcall raw_SetEOS () = 0;
	virtual HRESULT __stdcall raw_CopyTo (
		struct _Stream * DestStream,
		ADO_LONGPTR CharNumber) = 0;
	virtual HRESULT __stdcall raw_Flush () = 0;
	virtual HRESULT __stdcall raw_SaveToFile (
		BSTR FileName,
		enum SaveOptionsEnum Options) = 0;
	virtual HRESULT __stdcall raw_LoadFromFile (
		BSTR FileName) = 0;
	virtual HRESULT __stdcall raw_ReadText (
		long NumChars,
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall raw_WriteText (
		BSTR Data,
		enum StreamWriteEnum Options) = 0;
	virtual HRESULT __stdcall raw_Cancel () = 0;
};

struct __declspec(uuid("00000566-0000-0010-8000-00aa006d2ea4"))
Stream;
	// [ default ] interface _Stream

struct __declspec(uuid("00000567-0000-0010-8000-00aa006d2ea4"))
ADORecordConstruction : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetRow, put = PutRow))
	IUnknownPtr Row;
	__declspec(property(put = PutParentRow))
	IUnknownPtr ParentRow;

	//
	// Wrapper methods for error-handling
	//

	IUnknownPtr GetRow ();
	void PutRow (
		IUnknown * ppRow);
	void PutParentRow (
		IUnknown * _arg1);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Row (
		IUnknown * * ppRow) = 0;
	virtual HRESULT __stdcall put_Row (
		IUnknown * ppRow) = 0;
	virtual HRESULT __stdcall put_ParentRow (
		IUnknown * _arg1) = 0;
};

struct __declspec(uuid("00000568-0000-0010-8000-00aa006d2ea4"))
ADOStreamConstruction : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetStream, put = PutStream))
	IUnknownPtr Stream;

	//
	// Wrapper methods for error-handling
	//

	IUnknownPtr GetStream ();
	void PutStream (
		IUnknown * ppStm);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Stream (
		IUnknown * * ppStm) = 0;
	virtual HRESULT __stdcall put_Stream (
		IUnknown * ppStm) = 0;
};

struct __declspec(uuid("00000517-0000-0010-8000-00aa006d2ea4"))
ADOCommandConstruction : public IUnknown
{
	//
	// Property data
	//

	__declspec(property(get = GetOLEDBCommand, put = PutOLEDBCommand))
	IUnknownPtr OLEDBCommand;

	//
	// Wrapper methods for error-handling
	//

	IUnknownPtr GetOLEDBCommand ();
	void PutOLEDBCommand (
		IUnknown * ppOLEDBCommand);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_OLEDBCommand (
		IUnknown * * ppOLEDBCommand) = 0;
	virtual HRESULT __stdcall put_OLEDBCommand (
		IUnknown * ppOLEDBCommand) = 0;
};

struct __declspec(uuid("00000507-0000-0010-8000-00aa006d2ea4"))
Command;
	// [ default ] interface _Command

struct __declspec(uuid("00000535-0000-0010-8000-00aa006d2ea4"))
Recordset;
	// [ default ] interface _Recordset
	// [ default, source ] dispinterface RecordsetEvents

struct __declspec(uuid("00000283-0000-0010-8000-00aa006d2ea4"))
ADORecordsetConstruction : public IDispatch
{
	//
	// Property data
	//

	__declspec(property(get = GetRowset, put = PutRowset))
	IUnknownPtr Rowset;
	__declspec(property(get = GetChapter, put = PutChapter))
	ADO_LONGPTR Chapter;
	__declspec(property(get = GetRowPosition, put = PutRowPosition))
	IUnknownPtr RowPosition;

	//
	// Wrapper methods for error-handling
	//

	IUnknownPtr GetRowset ();
	void PutRowset (
		IUnknown * ppRowset);
	ADO_LONGPTR GetChapter ();
	void PutChapter (
		ADO_LONGPTR plChapter);
	IUnknownPtr GetRowPosition ();
	void PutRowPosition (
		IUnknown * ppRowPos);

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_Rowset (
		IUnknown * * ppRowset) = 0;
	virtual HRESULT __stdcall put_Rowset (
		IUnknown * ppRowset) = 0;
	virtual HRESULT __stdcall get_Chapter (
		ADO_LONGPTR * plChapter) = 0;
	virtual HRESULT __stdcall put_Chapter (
		ADO_LONGPTR plChapter) = 0;
	virtual HRESULT __stdcall get_RowPosition (
		IUnknown * * ppRowPos) = 0;
	virtual HRESULT __stdcall put_RowPosition (
		IUnknown * ppRowPos) = 0;
};

struct __declspec(uuid("00000505-0000-0010-8000-00aa006d2ea4"))
Field15 : public _ADO
{
	//
	// Property data
	//

	__declspec(property(get = GetValue, put = PutValue))
	_variant_t Value;
	__declspec(property(get = GetName))
	_bstr_t Name;
	__declspec(property(get = GetType))
	enum DataTypeEnum Type;
	__declspec(property(get = GetDefinedSize))
	ADO_LONGPTR DefinedSize;
	__declspec(property(get = GetOriginalValue))
	_variant_t OriginalValue;
	__declspec(property(get = GetUnderlyingValue))
	_variant_t UnderlyingValue;
	__declspec(property(get = GetActualSize))
	ADO_LONGPTR ActualSize;
	__declspec(property(get = GetPrecision))
	unsigned char Precision;
	__declspec(property(get = GetNumericScale))
	unsigned char NumericScale;
	__declspec(property(get = GetAttributes))
	long Attributes;

	//
	// Wrapper methods for error-handling
	//

	ADO_LONGPTR GetActualSize ();
	long GetAttributes ();
	ADO_LONGPTR GetDefinedSize ();
	_bstr_t GetName ();
	enum DataTypeEnum GetType ();
	_variant_t GetValue ();
	void PutValue (
		const _variant_t & pvar);
	unsigned char GetPrecision ();
	unsigned char GetNumericScale ();
	HRESULT AppendChunk (
		const _variant_t & Data);
	_variant_t GetChunk (
		long Length);
	_variant_t GetOriginalValue ();
	_variant_t GetUnderlyingValue ();

	//
	// Raw methods provided by interface
	//

	virtual HRESULT __stdcall get_ActualSize (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall get_Attributes (
		long * pl) = 0;
	virtual HRESULT __stdcall get_DefinedSize (
		ADO_LONGPTR * pl) = 0;
	virtual HRESULT __stdcall get_Name (
		BSTR * pbstr) = 0;
	virtual HRESULT __stdcall get_Type (
		enum DataTypeEnum * pDataType) = 0;
	virtual HRESULT __stdcall get_Value (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall put_Value (
		VARIANT pvar) = 0;
	virtual HRESULT __stdcall get_Precision (
		unsigned char * pbPrecision) = 0;
	virtual HRESULT __stdcall get_NumericScale (
		unsigned char * pbNumericScale) = 0;
	virtual HRESULT __stdcall raw_AppendChunk (
		VARIANT Data) = 0;
	virtual HRESULT __stdcall raw_GetChunk (
		long Length,
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_OriginalValue (
		VARIANT * pvar) = 0;
	virtual HRESULT __stdcall get_UnderlyingValue (
		VARIANT * pvar) = 0;
};

struct __declspec(uuid("0000050b-0000-0010-8000-00aa006d2ea4"))
Parameter;
	// [ default ] interface _Parameter


} // namespace XTPADODB

#include "XTPCalendarADO.inl"
//}}AFX_CODEJOCK_PRIVATE

#pragma pack(pop)

#endif // !defined(_XTPCALENDARADO_H__)
