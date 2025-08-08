package datetime

import (
	"testing"
)

func TestBetweenMonthsForDate(t *testing.T) {
	t.Log(BetweenMonthsForDate("2025-01", "2025-09", YYYYMM_0))
	t.Log(BetweenMonthsForDate("2024-07", "2025-03", YYYYMM_0))
	t.Log(BetweenMonthsForDate("2024-01", "2024-01", YYYYMM_0))
}
