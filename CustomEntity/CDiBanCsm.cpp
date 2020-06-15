#include "StdAfx.h"
#include "CDiBanCsm.h"


ACRX_DXF_DEFINE_MEMBERS(CDiBanCsm, AcDbEntity, AcDb::kDHL_CURRENT, AcDb::kMReleaseCurrent,
	AcDbProxyEntity::kNoOperation, CDiBanCsm, ZwCad);


CDiBanCsm::CDiBanCsm()
{
	m_arrEntPtr.append(&m_lineHori1);
	m_arrEntPtr.append(&m_lineHori2);
	m_arrEntPtr.append(&m_lineVert1);
	m_arrEntPtr.append(&m_lineVert2);

	m_vecGripName.clear();
	m_vecGripName.push_back(_T("BasePt"));
	m_vecGripName.push_back(_T("RightPt"));
	m_vecGripName.push_back(_T("TopPt"));
}


CDiBanCsm::CDiBanCsm(const AcGePoint3d& ptBase, double dLength, double dWidth)
	:m_ptBase(ptBase),m_dLength(dLength),m_dWidth(dWidth)
{
	//实体初始化
	m_ptRight = GetOffsetPt(m_ptBase, m_dLength, m_dWidth / 2);
	m_ptTop = GetOffsetPt(m_ptBase, m_dLength / 2, m_dWidth);

	m_lineHori1.setStartPoint(GetOffsetPt(m_ptBase,0, m_dWidth));
	m_lineHori1.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, m_dWidth));

	m_lineHori2.setStartPoint(m_ptBase);
	m_lineHori2.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, 0));

	m_lineVert1.setStartPoint(m_ptBase);
	m_lineVert1.setEndPoint(GetOffsetPt(m_ptBase, 0, m_dWidth));

	m_lineVert2.setStartPoint(GetOffsetPt(m_ptBase, m_dLength, 0));
	m_lineVert2.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, m_dWidth));
	//添加实体指针，用于调用基类
	m_arrEntPtr.append(&m_lineHori1);
	m_arrEntPtr.append(&m_lineHori2);
	m_arrEntPtr.append(&m_lineVert1);
	m_arrEntPtr.append(&m_lineVert2);
	
	//添加夹点类型
	m_vecGripName.clear();
	m_vecGripName.push_back(_T("BasePt"));
	m_vecGripName.push_back(_T("RightPt"));
	m_vecGripName.push_back(_T("TopPt"));
}

CDiBanCsm::CDiBanCsm(const CDiBanCsm& obj)
{
	m_lineHori1.setStartPoint(obj.m_lineHori1.startPoint());
	m_lineHori1.setEndPoint(obj.m_lineHori1.endPoint());

	m_lineHori2.setStartPoint(obj.m_lineHori2.startPoint());
	m_lineHori2.setEndPoint(obj.m_lineHori2.endPoint());

	m_lineVert1.setStartPoint(obj.m_lineVert1.startPoint());
	m_lineVert1.setEndPoint(obj.m_lineVert1.endPoint());

	m_lineVert2.setStartPoint(obj.m_lineVert2.startPoint());
	m_lineVert2.setEndPoint(obj.m_lineVert2.endPoint());

	m_ptBase = obj.m_ptBase;
	m_ptRight = obj.m_ptRight;
	m_ptTop = obj.m_ptTop;
	m_dLength = obj.m_dLength;
	m_dWidth = obj.m_dWidth;

	//添加实体指针，用于调用基类
	m_arrEntPtr.append(&m_lineHori1);
	m_arrEntPtr.append(&m_lineHori2);
	m_arrEntPtr.append(&m_lineVert1);
	m_arrEntPtr.append(&m_lineVert2);
	//添加夹点类型
	m_vecGripName.clear();
	m_vecGripName.push_back(_T("BasePt"));
	m_vecGripName.push_back(_T("RightPt"));
	m_vecGripName.push_back(_T("TopPt"));
}

CDiBanCsm::~CDiBanCsm()
{
}

Adesk::Boolean CDiBanCsm::worldDraw(AcGiWorldDraw* pWd)
{
	assertReadEnabled();
	return CCustomBase::worldDraw(pWd);
}

Acad::ErrorStatus CDiBanCsm::dwgInFields(AcDbDwgFiler* pFiler)
{//读取数据
	assertWriteEnabled();
	CCustomBase::dwgInFields(pFiler);
	pFiler->readPoint3d(&m_ptBase);
	pFiler->readPoint3d(&m_ptRight);
	pFiler->readPoint3d(&m_ptTop);

	return pFiler->filerStatus();
}

Acad::ErrorStatus CDiBanCsm::dwgOutFields(AcDbDwgFiler* pFiler) const
{//存入数据
	assertReadEnabled();
	CCustomBase::dwgOutFields(pFiler);
	pFiler->writePoint3d(m_ptBase);
	pFiler->writePoint3d(m_ptRight);
	pFiler->writePoint3d(m_ptTop);

	return pFiler->filerStatus();
}

Acad::ErrorStatus CDiBanCsm::dxfInFields(AcDbDxfFiler* pFiler)
{
	return Acad::eNotImplementedYet;
}

Acad::ErrorStatus CDiBanCsm::dxfOutFields(AcDbDxfFiler* pFiler) const
{
	return Acad::eNotImplementedYet;
}

Acad::ErrorStatus CDiBanCsm::getGripPoints(AcDbGripDataPtrArray& grips, const double curViewUnitSize, 
	const int gripSize, const AcGeVector3d& curViewDir, const int bitflags) const
{
	assertReadEnabled();
	std::vector<CString>::const_iterator appIter = GetGripName();
	AcDbGripDataPtrArray tempGripArr;

	//BasePt
	AcDbGripData* pGripBase = new AcDbGripData();
	pGripBase->setAppData((void*)&(*appIter));
	pGripBase->setGripPoint(m_ptBase);
	grips.append(pGripBase);
	tempGripArr.append(pGripBase);
	appIter++;

	//RightPt
	AcDbGripData* pGripRight = new AcDbGripData();
	pGripRight->setAppData((void*)&(*appIter));
	pGripRight->setGripPoint(m_ptRight);
	pGripRight->setWorldDraw(DrawStretchGrip);
	grips.append(pGripRight);
	tempGripArr.append(pGripRight);
	appIter++;

	//TopPt
	AcDbGripData* pGripTop = new AcDbGripData();
	pGripTop->setAppData((void*)&(*appIter));
	pGripTop->setGripPoint(m_ptTop);
	pGripTop->setWorldDraw(DrawStretchGrip);
	grips.append(pGripTop);
	tempGripArr.append(pGripTop);
	appIter++;

	auto a = this;
	CCustomBase::s_mapGripPtr[this] = tempGripArr;
	return Acad::eOk;
}

Acad::ErrorStatus CDiBanCsm::moveGripPointsAt(const AcDbVoidPtrArray& gripAppData,const AcGeVector3d& offset,
const int bitflags)
{
	assertWriteEnabled();
	CString* pStrTemp = NULL;
	for (int i = 0; i < gripAppData.length(); i++)
	{
		pStrTemp = static_cast<CString*>(gripAppData[i]);
		if (*pStrTemp == _T("BasePt"))
		{
			SetLinePoint(m_lineHori1, 3, offset);
			SetLinePoint(m_lineHori2, 3, offset);
			SetLinePoint(m_lineVert1, 3, offset);
			SetLinePoint(m_lineVert2, 3, offset);
			m_ptBase = GetBaseGripPt();
			m_ptRight = GetRightGripPt();
			m_ptTop = GetTopGripPt();
		}
		else if (*pStrTemp == _T("RightPt"))
		{
			AcGePoint3d ptStart = m_lineHori2.startPoint();
			AcGePoint3d ptEnd = m_lineHori2.endPoint();
			AcGeVector3d vctHori = ptEnd - ptStart;
			AcGeVector3d vctOffset = GetProjectVector(m_ptBase, offset, vctHori);
			SetLinePoint(m_lineHori1,2,vctOffset);
			SetLinePoint(m_lineHori2, 2, vctOffset);
			SetLinePoint(m_lineVert2, 3, vctOffset);
			m_ptBase = GetBaseGripPt();
			m_ptRight = GetRightGripPt();
			m_ptTop = GetTopGripPt();
		}
		else if (*pStrTemp == _T("TopPt"))
		{
			AcGePoint3d ptStart = m_lineVert1.startPoint();
			AcGePoint3d ptEnd = m_lineVert1.endPoint();
			AcGeVector3d vctHori = ptEnd - ptStart;
			AcGeVector3d vctOffset = GetProjectVector(m_ptBase, offset, vctHori);
			SetLinePoint(m_lineVert1, 2, vctOffset);
			SetLinePoint(m_lineVert2, 2, vctOffset);
			SetLinePoint(m_lineHori1, 3, vctOffset);
			m_ptBase = GetBaseGripPt();
			m_ptRight = GetRightGripPt();
			m_ptTop = GetTopGripPt();
		}
	}
	return Acad::eOk;
}

Acad::ErrorStatus CDiBanCsm::getOsnapPoints(AcDb::OsnapMode osnapMode, Adesk::GsMarker gsSelectionMark, 
	const AcGePoint3d& pickPoint, const AcGePoint3d& lastPoint, const AcGeMatrix3d& viewXform,
	AcGePoint3dArray& snapPoints, AcDbIntArray & geomIds) const
{
	assertReadEnabled();
	return CCustomBase::getOsnapPoints(osnapMode, gsSelectionMark, pickPoint, lastPoint, viewXform,
		snapPoints, geomIds);
}

Acad::ErrorStatus CDiBanCsm::getGeomExtents(AcDbExtents& extents) const
{
	assertReadEnabled();
	return CCustomBase::getGeomExtents(extents);
}

Acad::ErrorStatus CDiBanCsm::transformBy(const AcGeMatrix3d& xform)
{
	assertWriteEnabled();
	m_ptBase.transformBy(xform);
	m_ptRight.transformBy(xform);
	m_ptTop.transformBy(xform);

	return CCustomBase::transformBy(xform);
}

Acad::ErrorStatus CDiBanCsm::getTransformedCopy(const AcGeMatrix3d& xform, AcDbEntity*& ent) const
{
	assertReadEnabled();
	return CCustomBase::getTransformedCopy(xform, ent);
}

Acad::ErrorStatus CDiBanCsm::explode(ZcDbVoidPtrArray& entitySet) const
{
	assertReadEnabled();
// 	entitySet.append(m_lineHori1.clone());
// 	entitySet.append(m_lineHori2.clone());
// 	entitySet.append(m_lineVert1.clone());
// 	entitySet.append(m_lineVert2.clone());
	return Acad::eCannotExplodeEntity;
}

Acad::ErrorStatus CDiBanCsm::erase(Adesk::Boolean erasing /*= true*/)
{
	assertWriteEnabled();
	return CCustomBase::erase(erasing);
}

bool CDiBanCsm::SetPosition(const AcGePoint3d& pt)
{
	assertWriteEnabled();
	AcGeVector3d offset = pt - m_ptBase;
	SetLinePoint(m_lineHori1, 3, offset);
	SetLinePoint(m_lineHori2, 3, offset);
	SetLinePoint(m_lineVert1, 3, offset);
	SetLinePoint(m_lineVert2, 3, offset);
	m_ptBase = GetBaseGripPt();
	m_ptRight = GetRightGripPt();
	m_ptTop = GetTopGripPt();
	return true;
}

bool CDiBanCsm::SetRotation(double dAngle)
{
	assertWriteEnabled();
	m_lineHori1.setStartPoint(GetOffsetPt(m_ptBase, 0, m_dWidth).rotateBy(dAngle,AcGeVector3d::kZAxis,m_ptBase));
	m_lineHori1.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, m_dWidth).rotateBy(dAngle, AcGeVector3d::kZAxis, m_ptBase));
	m_lineHori2.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, 0).rotateBy(dAngle, AcGeVector3d::kZAxis, m_ptBase));
	m_lineVert1.setEndPoint(GetOffsetPt(m_ptBase, 0, m_dWidth).rotateBy(dAngle, AcGeVector3d::kZAxis, m_ptBase));
	m_lineVert2.setStartPoint(GetOffsetPt(m_ptBase, m_dLength, 0).rotateBy(dAngle, AcGeVector3d::kZAxis, m_ptBase));
	m_lineVert2.setEndPoint(GetOffsetPt(m_ptBase, m_dLength, m_dWidth).rotateBy(dAngle, AcGeVector3d::kZAxis, m_ptBase));
	m_ptBase = GetBaseGripPt();
	m_ptRight = GetRightGripPt();
	m_ptTop = GetTopGripPt();
	return true;
}

AcGeVector3d CDiBanCsm::GetHoriDirection() const
{
	AcGePoint3d ptStart = m_lineHori2.startPoint();
	AcGePoint3d ptEnd = m_lineHori2.endPoint();
	return ptEnd - ptStart;
}

AcGeVector3d CDiBanCsm::GetVertDirection() const
{
	AcGePoint3d ptStart = m_lineVert1.startPoint();
	AcGePoint3d ptEnd = m_lineVert1.endPoint();
	return ptEnd - ptStart;
}

void CDiBanCsm::gripStatus(const AcDb::GripStat status)
{
	AcDbEntity::gripStatus(status);

	switch (status)
	{
	case AcDb::kGripsDone:
		CCustomBase::DeleteGrip(this);
		break;
	case AcDb::kGripsToBeDeleted:
		
		break;
	case AcDb::kDimDataToBeDeleted:

		break;
	}
}

AcGePoint3d CDiBanCsm::GetBaseGripPt() const
{
	return m_lineHori2.startPoint();
}

AcGePoint3d CDiBanCsm::GetRightGripPt() const
{
	AcGePoint3d ptStart = m_lineVert2.startPoint();
	AcGePoint3d ptEnd = m_lineVert2.endPoint();
	return AcGePoint3d((ptStart.x + ptEnd.x)/2,(ptStart.y + ptEnd.y)/2,(ptStart.z + ptEnd.z)/2);
}

AcGePoint3d CDiBanCsm::GetTopGripPt() const
{
	AcGePoint3d ptStart = m_lineHori1.startPoint();
	AcGePoint3d ptEnd = m_lineHori1.endPoint();
	return AcGePoint3d((ptStart.x + ptEnd.x) / 2, (ptStart.y + ptEnd.y) / 2, (ptStart.z + ptEnd.z) / 2);
}

bool CDiBanCsm::DrawStretchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize)
{
	dGripSize *= CCustomBase::s_gripSize / 2;
	AcGePoint3d	pt = pThis->gripPoint();
	CString* pStr = (CString*)(pThis->appData());
	AcGeVector3d vec = AcGeVector3d::kXAxis;
	AcDbObjectPointer<CDiBanCsm> m_pDiBan(entId, AcDb::kForRead);
	if (m_pDiBan.openStatus() != Acad::eOk)
		return false;

	if (*pStr == _T("RightPt"))
		vec = m_pDiBan->GetHoriDirection().normal();
	else if (*pStr == _T("TopPt"))
		vec = m_pDiBan->GetVertDirection().normal();

	auto vecUp = AcGeVector3d::kZAxis.crossProduct(vec).normal();
	AcGePoint3d pts[3];
	pts[0] = pt + vec * dGripSize;
	pts[1] = pt - vec * dGripSize / 2 + vecUp * dGripSize * .7;
	pts[2] = pt - vec * dGripSize / 2 - vecUp * dGripSize * .7;

	pWd->subEntityTraits().setFillType(kAcGiFillAlways);
	pWd->geometry().polygon(3, pts);

	return true;
}

bool CDiBanCsm::DrawSwitchGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize)
{
	dGripSize *= CCustomBase::s_gripSize / 2;
	AcGePoint3d	pt = pThis->gripPoint();
	AcGeVector3d vec = AcGeVector3d::kXAxis;

	auto vecUp = AcGeVector3d::kZAxis.crossProduct(vec).normal();
	AcGePoint3d ptsRec[4];
	ptsRec[0] = pt + vecUp * dGripSize;
	ptsRec[1] = pt + vecUp * dGripSize - vec * dGripSize;
	ptsRec[2] = pt - vecUp * dGripSize - vec * dGripSize;
	ptsRec[3] = pt - vecUp * dGripSize;

	AcGePoint3d ptsArr[3];
	ptsArr[0] = pt + vec * dGripSize * 2.4;
	ptsArr[1] = pt + vec * dGripSize + vecUp * dGripSize;
	ptsArr[2] = pt + vec * dGripSize - vecUp * dGripSize;

	pWd->subEntityTraits().setFillType(kAcGiFillAlways);

	pWd->geometry().polygon(4, ptsRec);
	pWd->geometry().polygon(3, ptsArr);
	return true;
}

bool CDiBanCsm::DrawEnumGrip(AcDbGripData *pThis, AcGiWorldDraw *pWd, const AcDbObjectId& entId, AcDbGripOperations::DrawType type, AcGePoint3d *cursor, double dGripSize)
{
	dGripSize *= CCustomBase::s_gripSize / 2;
	AcGePoint3d	pt = pThis->gripPoint();

	// 中心圆
	pWd->subEntityTraits().setFillType(kAcGiFillAlways);
	pWd->geometry().circle(pt, dGripSize, AcGeVector3d::kZAxis);

	return true;
}

