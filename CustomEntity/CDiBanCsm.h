#pragma once
#include "CCustomBase.h"

//底板
class CDiBanCsm :public CCustomBase
{
public:
	ACRX_DECLARE_MEMBERS(CDiBanCsm);

	CDiBanCsm();
	CDiBanCsm(const AcGePoint3d& ptBase,double dLength,double dWidth);
	CDiBanCsm(const CDiBanCsm& obj);//拷贝构造函数
	~CDiBanCsm();

#pragma region 重载函数
	virtual Acad::ErrorStatus getGripPoints(AcDbGripDataPtrArray& grips, const double curViewUnitSize, const int gripSize, const AcGeVector3d& curViewDir, const int bitflags) const;
	virtual Acad::ErrorStatus moveGripPointsAt(const AcDbVoidPtrArray& gripAppData,const AcGeVector3d& offset,const int bitflags);
	virtual Acad::ErrorStatus dwgInFields(AcDbDwgFiler* pFiler);//加载
	virtual Acad::ErrorStatus dwgOutFields(AcDbDwgFiler* pFiler) const;//保存
	virtual Acad::ErrorStatus transformBy(const AcGeMatrix3d& xform);
	virtual Acad::ErrorStatus explode(ZcDbVoidPtrArray& entitySet) const;
	virtual void gripStatus(const AcDb::GripStat status);

	virtual Adesk::Boolean worldDraw(AcGiWorldDraw* pWd);//绘制
	virtual Acad::ErrorStatus dxfInFields(AcDbDxfFiler* pFiler);
	virtual Acad::ErrorStatus dxfOutFields(AcDbDxfFiler* pFiler) const;
	virtual Acad::ErrorStatus getOsnapPoints(AcDb::OsnapMode osnapMode, Adesk::GsMarker gsSelectionMark, const AcGePoint3d& pickPoint, const AcGePoint3d& lastPoint,
		const AcGeMatrix3d& viewXform, AcGePoint3dArray& snapPoints, AcDbIntArray & geomIds) const;
	virtual Acad::ErrorStatus getGeomExtents(AcDbExtents& extents) const;
	virtual Acad::ErrorStatus getTransformedCopy(const AcGeMatrix3d& xform, AcDbEntity*& ent) const;
	virtual	Acad::ErrorStatus erase(Adesk::Boolean erasing = true);
	
#pragma endregion 重载函数

#pragma region 接口
	bool SetPosition(const AcGePoint3d& pt);//设置基点位置
	bool SetRotation(double dAngle);//设置旋转角度
	AcGeVector3d GetHoriDirection() const;//得到水平向量
	AcGeVector3d GetVertDirection() const;//得到垂直向量
#pragma endregion 接口

#pragma region 辅助函数
protected:
	//得到夹点坐标
	AcGePoint3d GetBaseGripPt() const;
	AcGePoint3d GetRightGripPt() const;
	AcGePoint3d GetTopGripPt() const;
#pragma endregion 辅助函数

#pragma region 夹点
public:
	//拉伸夹点
	static bool DrawStretchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
	//转换夹点
	static bool DrawSwitchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
	//枚举夹点
	static bool DrawEnumGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize);
#pragma endregion 夹点

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

