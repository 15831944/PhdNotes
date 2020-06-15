#pragma once

/*
//�Զ���ʵ���ʼ��
CDiBanCsm::rxInit();
acrxBuildClassHierarchy();
//�Զ���ʵ��ɾ��
deleteAcRxClass(CDiBanCsm::desc());
*/

//�Զ���ʵ�����
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
#pragma region ��������
	std::vector<CString>::const_iterator GetGripName() const;//��¼�е������

	//************************************
	// Summary: 	�õ�X���Y��ƫ�Ƶ��ĵ�����
	// Parameters: 	
	//    ptBase		- 	����Ҫƫ�Ƶĵ�����
	//    dX		- 	����X��ƫ�ƾ���
	//    dY		- 	����Y��ƫ�ƾ���
	// Returns:   	
	//************************************
	AcGePoint3d GetOffsetPt(const AcGePoint3d& ptBase, double dX, double dY);

	//************************************
	// Summary: 	�õ�ĳһ������ָ�������ϵ�ͶӰ����
	// Parameters: 	
	//    ptBase		- 	�������������Ľ���
	//    vctOffset		- 	����ĳһ����
	//    vctDirection		- 	����ָ������
	// Returns:   	
	//************************************
	AcGeVector3d GetProjectVector(const AcGePoint3d& ptBase, const AcGeVector3d& vctOffset, const AcGeVector3d& vctDirection);
	
	//************************************
	// Summary: 	ͨ��ƫ�������ƶ�ֱ�ߵĵ�����
	// Parameters: 	
	//    pLine		- 	����ֱ��ָ��
	//    nType		- 	�������ͣ�1���ƶ���㣬2���ƶ��յ㣬3�����ƶ���
	//    vctOffset		- 	����ƫ������
	// Returns:   	
	//************************************
	void SetLinePoint(AcDbLine& line, int nType, const AcGeVector3d& vctOffset);//����ֱ��ƫ�Ƶ�
	
#pragma endregion ��������

#pragma region �е�
public:
	static double s_gripSize;	//�е�Ĵ�С
	static std::map<const AcDbEntity*, AcDbGripDataPtrArray> s_mapGripPtr;

	//delete�е�
	static bool DeleteGrip(const AcDbEntity* pEnt);
#pragma endregion �е�

protected:
	AcArray<AcDbEntity*> m_arrEntPtr;
	std::vector<CString> m_vecGripName;
};


