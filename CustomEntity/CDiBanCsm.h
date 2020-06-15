#pragma once
#include "CCustomBase.h"

//�װ�
class CDiBanCsm :public CCustomBase
{
public:
	ACRX_DECLARE_MEMBERS(CDiBanCsm);

	CDiBanCsm();
	CDiBanCsm(const AcGePoint3d& ptBase,double dLength,double dWidth);
	CDiBanCsm(const CDiBanCsm& obj);//�������캯��
	~CDiBanCsm();

#pragma region ���غ���
	virtual Acad::ErrorStatus getGripPoints(AcDbGripDataPtrArray& grips, const double curViewUnitSize, const int gripSize, const AcGeVector3d& curViewDir, const int bitflags) const;
	virtual Acad::ErrorStatus moveGripPointsAt(const AcDbVoidPtrArray& gripAppData,const AcGeVector3d& offset,const int bitflags);
	virtual Acad::ErrorStatus dwgInFields(AcDbDwgFiler* pFiler);//����
	virtual Acad::ErrorStatus dwgOutFields(AcDbDwgFiler* pFiler) const;//����
	virtual Acad::ErrorStatus transformBy(const AcGeMatrix3d& xform);
	virtual Acad::ErrorStatus explode(ZcDbVoidPtrArray& entitySet) const;
	virtual void gripStatus(const AcDb::GripStat status);

	virtual Adesk::Boolean worldDraw(AcGiWorldDraw* pWd);//����
	virtual Acad::ErrorStatus dxfInFields(AcDbDxfFiler* pFiler);
	virtual Acad::ErrorStatus dxfOutFields(AcDbDxfFiler* pFiler) const;
	virtual Acad::ErrorStatus getOsnapPoints(AcDb::OsnapMode osnapMode, Adesk::GsMarker gsSelectionMark, const AcGePoint3d& pickPoint, const AcGePoint3d& lastPoint,
		const AcGeMatrix3d& viewXform, AcGePoint3dArray& snapPoints, AcDbIntArray & geomIds) const;
	virtual Acad::ErrorStatus getGeomExtents(AcDbExtents& extents) const;
	virtual Acad::ErrorStatus getTransformedCopy(const AcGeMatrix3d& xform, AcDbEntity*& ent) const;
	virtual	Acad::ErrorStatus erase(Adesk::Boolean erasing = true);
	
#pragma endregion ���غ���

#pragma region �ӿ�
	bool SetPosition(const AcGePoint3d& pt);//���û���λ��
	bool SetRotation(double dAngle);//������ת�Ƕ�
	AcGeVector3d GetHoriDirection() const;//�õ�ˮƽ����
	AcGeVector3d GetVertDirection() const;//�õ���ֱ����
#pragma endregion �ӿ�

#pragma region ��������
protected:
	//�õ��е�����
	AcGePoint3d GetBaseGripPt() const;
	AcGePoint3d GetRightGripPt() const;
	AcGePoint3d GetTopGripPt() const;
#pragma endregion ��������

#pragma region �е�
public:
	//����е�
	static bool DrawStretchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
	//ת���е�
	static bool DrawSwitchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
	//ö�ټе�
	static bool DrawEnumGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
#pragma endregion �е�

protected:
	AcDbLine m_lineHori1;
	AcDbLine m_lineHori2;
	AcDbLine m_lineVert1;
	AcDbLine m_lineVert2;

	AcGePoint3d m_ptBase;	
	AcGePoint3d m_ptRight;	
	AcGePoint3d m_ptTop;
	double m_dLength;
	double m_dWidth;

};

