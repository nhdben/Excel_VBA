Function payback_period(init_investment, cinflow)
  'c = Rng.Columns.Count
  'r = Rng.Rows.Count
  x = Abs(init_investment)
  i = 0
  j = cinflow.Count
  'If c > 1 And r > 1 Then
   payback_period = "Error"
  'Else
  Do While i < j
      i = i + 1
      x = x - y
      y = cinflow.Cells(i).Value
      If x = y Then
          payback_period = i
          Exit Function
      ElseIf x < y Then
          prev_year = i - 1
          frac_year = x / y
          payback_period = prev_year + frac_year
          Exit Function
      End If
  Loop
  payback_period = "Project does not pay back"
  'End If
  End Function
