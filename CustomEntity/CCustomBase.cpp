#include "StdAfx.h"
#include "CCustomBase.h"
#include <strsafe.h>
#include "../NameSpace/PhdInline.h"

double CCustomBase::s_gripSize = 2.8;
std::map<const AcDbEntity*, AcDbGripDataPtrArray> CCustomBase::s_mapGripPtr;

CCustomBase::CCustomBase()
{
}


CCustomBase::~CCustomBase()
{
}

Adesk::Boolean CCustomBase::worldDraw(AcGiWorldDraw* pWd)
{
	assertReadEnabled();
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->worldDraw(pWd);
	}
	return Adesk::kTrue;
}

Acad::ErrorStatus CCustomBase::dwgInFields(AcDbDwgFiler* pFiler)
{
	//读取数据
	assertWriteEnabled();
	AcDbEntity::dwgInFields(pFiler);
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->dwgInFields(pFiler);
	}
	return pFiler->filerStatus();
}

Acad::ErrorStatus CCustomBase::dwgOutFields(AcDbDwgFiler* pFiler) const
{
	//存入数据
	assertReadEnabled();
	AcDbEntity::dwgOutFields(pFiler);
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->dwgOutFields(pFiler);
	}
	return pFiler->filerStatus();
}

Acad::ErrorStatus CCustomBase::dxfInFields(AcDbDxfFiler* pFiler)
{
	return Acad::eNotImplementedYet;
}

Acad::ErrorStatus CCustomBase::dxfOutFields(AcDbDxfFiler* pFiler) const
{
	return Acad::eNotImplementedYet;
}

Acad::ErrorStatus CCustomBase::getOsnapPoints(AcDb::OsnapMode osnapMode, Adesk::GsMarker gsSelectionMark, const AcGePoint3d& pickPoint, const AcGePoint3d& lastPoint, const AcGeMatrix3d& viewXform, AcGePoint3dArray& snapPoints, AcDbIntArray & geomIds) const
{
	assertReadEnabled();
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->getOsnapPoints(osnapMode, gsSelectionMark, pickPoint, lastPoint, viewXform, snapPoints, geomIds);
	}
	return Acad::eOk;
}

Acad::ErrorStatus CCustomBase::getGeomExtents(AcDbExtents& extents) const
{
	assertReadEnabled();
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		AcDbExtents ext;
		m_arrEntPtr[i]->getGeomExtents(ext);
		extents.addExt(ext);
	}
	return Acad::eOk;
}

Acad::ErrorStatus CCustomBase::transformBy(const AcGeMatrix3d& xform)
{
	assertWriteEnabled();
	
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->transformBy(xform);
	}
	return AcDbEntity::transformBy(xform);
}

Acad::ErrorStatus CCustomBase::getTransformedCopy(const AcGeMatrix3d& xform, AcDbEntity*& ent) const
{
	assertReadEnabled();
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->getTransformedCopy(xform, ent);
	}
	return AcDbEntity::getTransformedCopy(xform, ent);
}

Acad::ErrorStatus CCustomBase::erase(Adesk::Boolean erasing /*= true*/)
{
	assertWriteEnabled();
	for (int i = 0; i < m_arrEntPtr.length(); ++i)
	{
		m_arrEntPtr[i]->erase(erasing);
	}
	return Acad::eOk;
}

std::vector<CString>::const_iterator CCustomBase::GetGripName() const
{
	return m_vecGripName.begin();
}

AcGePoint3d CCustomBase::GetOffsetPt(const AcGePoint3d& ptBase, double dX, double dY)
{
	AcGePoint3d ptTemp = ptBase;
	ptTemp.x += dX;
	ptTemp.y += dY;
	return ptTemp;
}

AcGeVector3d CCustomBase::GetProjectVector(const AcGePoint3d& ptBase, const AcGeVector3d& vctOffset, const AcGeVector3d& vctDirection)
{
	AcGePoint3d ptOffset = ptBase + vctOffset;
	AcGeLine3d geLine(ptBase, vctDirection);
	AcGePoint3d ptTarget = geLine.closestPointTo(ptOffset);
	return ptTarget - ptBase;
}

void CCustomBase::SetLinePoint(AcDbLine& line, int nType, const AcGeVector3d& vctOffset)
{
	AcGePoint3d ptStart = line.startPoint();
	AcGePoint3d ptEnd = line.endPoint();
	switch (nType)
	{
	case 1://移动起点
		ptStart += vctOffset;
		line.setStartPoint(ptStart);
		break;
	case 2://移动终点
		ptEnd += vctOffset;
		line.setEndPoint(ptEnd);
		break;
	case 3://移动起点和终点
		ptStart += vctOffset;
		line.setStartPoint(ptStart);
		ptEnd += vctOffset;
		line.setEndPoint(ptEnd);
		break;
	default:
		break;
	}
}

bool CCustomBase::DeleteGrip(const AcDbEntity* pEnt)
{
	std::map<const AcDbEntity*, AcDbGripDataPtrArray>::iterator iter;
	for (iter = s_mapGripPtr.begin(); iter != s_mapGripPtr.end(); iter++)
	{
		if (pEnt == iter->first)
		{
			AcDbGripDataPtrArray tempGripArr = iter->second;
			for (int i = 0; i < tempGripArr.length(); ++i)
			{
				AcDbGripData* pGrip = tempGripArr[i];
				DEL(tempGripArr[i]);
				tempGripArr[i] = NULL;
			}
		}
	}
	int n = CCustomBase::s_mapGripPtr.erase(pEnt);//如果删除了会返回1，否则返回0  
	return n;
}



