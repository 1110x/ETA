using System.Collections.Generic;
using ETA.Models;

namespace ETA.Services.SERVICE1;

/// <summary>분석단가 테이블 제거로 인해 빈 스텁으로 유지. 단가는 계약 DB 컬럼에서 직접 조회.</summary>
public static class ContractPriceService
{
    public static List<ContractPrice> GetAllContractPrices() => new();
}
