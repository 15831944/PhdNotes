#pragma once

/*
//自定义实体初始化
CDiBanCsm::rxInit();
acrxBuildClassHierarchy();
//自定义实体删除
deleteAcRxClass(CDiBanCsm::desc());
*/

//自定义实体基类
class CCustomBase :public AcDbEntity
{
public:
	CCustomBase();
	~CCustomBase();

	virtual Adesk::Boolean worldDraw(AcGiWorldDraw* pWd);
	virtual Acad::ErrorStatus dwgInFields(AcDbDwgFiler* pFiler);
	virtual Acad::ErrorStatus dwgOutFields(AcDbDwgFiler* pFiler) const;
	virtual Acad::ErrorStatus dxfInFields(AcDbDxfFiler* pFiler);
	virtual Acad::ErrorStatus dxfOutFields(AcDbDxfFiler* pFiler) const;
	virtual Acad::ErrorStatus getOsnapPoints(AcDb::OsnapMode osnapMode, Adesk::GsMarker gsSelectionMark, const AcGePoint3d& pickPoint, const AcGePoint3d& lastPoint,
		const AcGeMatrix3d& viewXform, AcGePoint3dArray& snapPoints, AcDbIntArray & geomIds) const;
	virtual Acad::ErrorStatus getGeomExtents(AcDbExtents& extents) const;
	virtual Acad::ErrorStatus transformBy(const AcGeMatrix3d& xform);
	virtual Acad::ErrorStatus getTransformedCopy(const AcGeMatrix3d& xform, AcDbEntity*& ent) const;
	virtual	Acad::ErrorStatus erase(Adesk::Boolean erasing = true);

protected:
#pragma region 辅助函数
	std::vector<CString>::const_iterator GetGripName() const;//记录夹点的名称

	//************************************
	// Summary: 	得到X轴或Y轴偏移点后的点坐标
	// Parameters: 	
	//    ptBase		- 	输入要偏移的点坐标
	//    dX		- 	输入X轴偏移距离
	//    dY		- 	输入Y轴偏移距离
	// Returns:   	
	//************************************
	AcGePoint3d GetOffsetPt(const AcGePoint3d& ptBase, double dX, double dY);

	//************************************
	// Summary: 	得到某一向量在指定向量上的投影向量
	// Parameters: 	
	//    ptBase		- 	输入两个向量的交点
	//    vctOffset		- 	输入某一向量
	//    vctDirection		- 	输入指定向量
	// Returns:   	
	//************************************
	AcGeVector3d GetProjectVector(const AcGePoint3d& ptBase, const AcGeVector3d& vctOffset, const AcGeVector3d& vctDirection);
	
	//************************************
	// Summary: 	通过偏移向量移动直线的点坐标
	// Parameters: 	
	//    pLine		- 	输入直线指针
	//    nType		- 	输入类型（1：移动起点，2：移动终点，3：都移动）
	//    vctOffset		- 	输入偏移向量
	// Returns:   	
	//************************************
	void SetLinePoint(AcDbLine& line, int nType, const AcGeVector3d& vctOffset);//设置直线偏移点
	
#pragma endregion 辅助函数

#pragma region 夹点
public:
	static double s_gripSize;	//夹点的大小
	static std::map<const AcDbEntity*, AcDbGripDataPtrArray> s_mapGripPtr;

	//delete夹点
	static bool DeleteGrip(const AcDbEntity* pEnt);
#pragma endregion 夹点

protected:
	AcArray<AcDbEntity*> m_arrEntPtr;
	std::vector<CString> m_vecGripName;
};


