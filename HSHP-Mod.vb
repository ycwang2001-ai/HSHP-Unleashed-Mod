Imports BepInEx
Imports HarmonyLib
Imports zip.lexy.tgame.state.city
Imports zip.lexy.tgame.state.city.mayor
Imports zip.lexy.tgame.simulation
Imports zip.lexy.tgame.constants
Imports zip.lexy.tgame.ui.widget.build
Imports zip.lexy.tgame.state.building
Imports zip.lexy.tgame.city.building
Imports zip.lexy.tgame.city.building.details
Imports UnityEngine
Imports UnityEngine.UI
Imports System.Reflection
Imports System.Collections.Generic
Imports zip.lexy.tgame.state
Imports zip.lexy.tgame.state.trader
Imports zip.lexy.tgame.simulation.trading
Imports zip.lexy.tgame.localization
Imports System.Linq
Imports zip.lexy.tgame.state.pop
Imports zip.lexy.tgame.ui.widget.trader
Imports zip.lexy.tgame.ui.happiness
Imports zip.lexy.tgame.goods
Imports TMPro
Imports zip.lexy.tgame.ui.widget.trade
Imports zip.lexy.tgame.ui.widget.trader.townhall
Imports zip.lexy.tgame.state.ship
Imports zip.lexy.tgame.ui.business
Imports zip.lexy.tgame.simulation.consumption
Imports zip.lexy.tgame.ui.newgame
Imports UnityEngine.Events
Imports zip.lexy.tgame.ui.general
Imports zip.lexy.tgame.ui.quest
Imports zip.lexy.tgame.state.quest
Imports zip.lexy.tgame.quest

<BepInPlugin("com.yourname.hshpmod", "HSHP Mod", "1.0.0")>
Public Class HSHPModPlugin
    Inherits BaseUnityPlugin

    Private Sub Awake()
        Logger.LogInfo("HSHP Mod is loaded!")

        ' Initialize Harmony patches
        Dim harmony As New Harmony("com.yourname.hshpmod")
        harmony.PatchAll()

        ' Init HSHP quest registrar (registers custom QuestSO in atlas)
        HSHPQuestRegistrar.Init()
    End Sub

End Class

' Patch to delay mayor elections indefinitely
<HarmonyPatch(GetType(Mayor), "GetElectionTurn", MethodType.Normal)>
Public Class MayorElectionPatch
    Public Shared Function Prefix(ByRef __result As Integer) As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        ' Set election turn to 9999 (effectively never)
        __result = 9999
        ' Return False to skip the original method
        Return False
    End Function
End Class

' Patch to prevent automatic mayor quest filling
<HarmonyPatch(GetType(CityTurnAdvanceManager), "AutomaticallyFillNextMayorQuestIfNeeded", MethodType.Normal)>
Public Class PreventAutoMayorQuestPatch
    Public Shared Function Prefix() As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        ' Return False to skip the method entirely
        Return False
    End Function
End Class

' Patch to show/hide filter buttons when build window opens.
' FB's Changes only: hide everything except All.
' BuildMode only: rename raw->Businesses, processed->Shops, food->Dwellings; show all + bonus.
' Both FB's Changes + BuildMode: same as BuildMode only.
' Ship button is always hidden when either is active.
<HarmonyPatch(GetType(BuildWindow), "OnOpenBuildMenu", MethodType.Normal)>
Public Class BuildWindowOpenPatch
    Private Shared ReadOnly RenamedFilters As New Dictionary(Of String, String) From {
        {"raw", "Businesses"},
        {"processed", "Shops"},
        {"food", "Dwellings"}
    }

    Public Shared Sub Postfix(ByVal __instance As BuildWindow)
        Dim personalOn = HSHPPersonal.IsEnabled()
        Dim buildOn = HSHPBuildMode.IsEnabled()
        If Not personalOn AndAlso Not buildOn Then Return

        Dim allTransforms = __instance.gameObject.GetComponentsInChildren(Of Transform)(True)

        For Each t In allTransforms
            Dim objName = t.gameObject.name
            Dim nameLower = objName.ToLower()

            ' Ensure build-filters container is visible
            If objName.Equals("build-filters") Then
                t.gameObject.SetActive(True)
                Continue For
            End If

            If t.parent Is Nothing OrElse Not t.parent.name.ToLower().Contains("filter") Then Continue For

            ' "all" and "bonus" buttons: always visible
            If nameLower = "all" OrElse nameLower = "bonus" Then
                t.gameObject.SetActive(True)
                Continue For
            End If

            ' raw/processed/food buttons
            If RenamedFilters.ContainsKey(nameLower) Then
                If buildOn Then
                    ' BuildMode on: show renamed
                    t.gameObject.SetActive(True)
                    Dim tmpText = t.gameObject.GetComponentInChildren(Of TMP_Text)(True)
                    If tmpText IsNot Nothing Then tmpText.text = RenamedFilters(nameLower)
                Else
                    ' FB's Changes only (no BuildMode): hide these
                    t.gameObject.SetActive(False)
                End If
                Continue For
            End If

            ' Hide ship button
            If nameLower = "ship" Then
                t.gameObject.SetActive(False)
            End If
        Next
    End Sub
End Class

' Patch FilterAll to show buildable buildings when build mode or FB's Changes on.
' When HSHPPersonal is enabled, only the city's bonus goods are shown (+ shops/dwellings if build mode).
' When only HSHPBuildMode is enabled, all 21 goods + shops + dwellings are shown.
<HarmonyPatch(GetType(BuildWindow), "FilterAll", MethodType.Normal)>
Public Class FilterAllPatch
    Public Shared Function Prefix(ByVal __instance As BuildWindow) As Boolean
        If HSHPPersonal.IsEnabled() Then
            Dim gameState = InstanceProvider.GetInstance(Of GameState)()
            Dim goods As New List(Of String)(gameState.viewCity.bonuses)
            If HSHPBuildMode.IsEnabled() Then
                goods.AddRange({"shop-bakery", "shop-tavern", "shop-butcher",
                    "shop-blacksmith", "shop-cooks-haven", "shop-tailor", "shop-carpenter"})
                goods.AddRange({"dwelling-low", "dwelling-mid", "dwelling-high"})
            End If
            __instance.ShowBusinessesForGoods(goods)
            Return False
        End If

        If Not HSHPBuildMode.IsEnabled() Then Return True

        Dim allGoods As New List(Of String) From {
            "ale", "beef", "bricks", "clay", "cloth", "fish", "grain", "honey",
            "iron-bars", "logs", "lumber", "mead", "ore", "pottery", "salt",
            "stone", "tar", "vegetables", "wine", "wooden-tools", "wool"}
        allGoods.AddRange({"shop-bakery", "shop-tavern", "shop-butcher",
            "shop-blacksmith", "shop-cooks-haven", "shop-tailor", "shop-carpenter"})
        allGoods.AddRange({"dwelling-low", "dwelling-mid", "dwelling-high"})
        __instance.ShowBusinessesForGoods(allGoods)
        Return False
    End Function
End Class

' Patch FilterRaw to show only production goods (button label: "Businesses") when build mode on.
' When FB's Changes is enabled, only the city's bonus goods are shown.
<HarmonyPatch(GetType(BuildWindow), "FilterRaw", MethodType.Normal)>
Public Class FilterRawPatch
    Public Shared Function Prefix(ByVal __instance As BuildWindow) As Boolean
        If HSHPPersonal.IsEnabled() AndAlso HSHPBuildMode.IsEnabled() Then
            Dim gameState = InstanceProvider.GetInstance(Of GameState)()
            Dim goods As New List(Of String)(gameState.viewCity.bonuses)
            __instance.ShowBusinessesForGoods(goods)
            Return False
        End If

        If Not HSHPBuildMode.IsEnabled() Then Return True

        Dim allGoods As New List(Of String) From {
            "ale", "beef", "bricks", "clay", "cloth", "fish", "grain", "honey",
            "iron-bars", "logs", "lumber", "mead", "ore", "pottery", "salt",
            "stone", "tar", "vegetables", "wine", "wooden-tools", "wool"}
        __instance.ShowBusinessesForGoods(allGoods)
        Return False
    End Function
End Class

' Patch FilterProcessed to show only shops (button label: "Shops") when build mode on
<HarmonyPatch(GetType(BuildWindow), "FilterProcessed", MethodType.Normal)>
Public Class FilterProcessedPatch
    Public Shared Function Prefix(ByVal __instance As BuildWindow) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True

        Dim goods As New List(Of String) From {
            "shop-bakery", "shop-tavern", "shop-butcher",
            "shop-blacksmith", "shop-cooks-haven", "shop-tailor", "shop-carpenter"
        }
        __instance.ShowBusinessesForGoods(goods)
        Return False
    End Function
End Class

' Patch FilterFood to show only dwellings (button label: "Dwellings") when build mode on
<HarmonyPatch(GetType(BuildWindow), "FilterFood", MethodType.Normal)>
Public Class FilterFoodPatch
    Public Shared Function Prefix(ByVal __instance As BuildWindow) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True

        Dim goods As New List(Of String) From {
            "dwelling-low", "dwelling-mid", "dwelling-high"
        }
        __instance.ShowBusinessesForGoods(goods)
        Return False
    End Function
End Class

' Patch to lower reputation level requirements
<HarmonyPatch(GetType(Reputation), "GetReputationLevelByPoints", MethodType.Normal)>
Public Class ReputationLevelPatch
    Public Shared Function Prefix(ByVal points As Single, ByRef __result As String) As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        ' Custom thresholds: 0, 50, 100, 250, 500, 1000, 2000, 5000, 10000, 20000
        If points < 1000.0F Then
            If points < 250.0F Then
                If points < 50.0F Then
                    __result = "unknown"
                ElseIf points < 100.0F Then
                    __result = "obscure"
                Else
                    __result = "emerging"
                End If
            ElseIf points < 500.0F Then
                __result = "recognized"
            Else
                __result = "promising"
            End If
        ElseIf points < 5000.0F Then
            If points < 2000.0F Then
                __result = "known"
            Else
                __result = "respected"
            End If
        ElseIf points < 10000.0F Then
            __result = "honored"
        ElseIf points < 20000.0F Then
            __result = "revered"
        Else
            __result = "renowned"
        End If

        Return False ' Skip original method
    End Function
End Class

' Patch to lower next reputation points requirements
<HarmonyPatch(GetType(Reputation), "GetNextReputationPoints", MethodType.Normal)>
Public Class NextReputationPointsPatch
    Public Shared Function Prefix(ByVal points As Single, ByRef __result As Long) As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        ' Custom thresholds: 0, 50, 100, 250, 500, 1000, 2000, 5000, 10000, 20000
        If points < 1000.0F Then
            If points < 250.0F Then
                If points < 50.0F Then
                    __result = 50L
                ElseIf points < 100.0F Then
                    __result = 100L
                Else
                    __result = 250L
                End If
            ElseIf points < 500.0F Then
                __result = 500L
            Else
                __result = 1000L
            End If
        ElseIf points < 5000.0F Then
            If points < 2000.0F Then
                __result = 2000L
            Else
                __result = 5000L
            End If
        ElseIf points < 10000.0F Then
            __result = 10000L
        ElseIf points < 20000.0F Then
            __result = 20000L
        Else
            __result = -1L ' Max level reached
        End If

        Return False ' Skip original method
    End Function
End Class

' Patch to give more reputation and mayor respect when selling goods
<HarmonyPatch(GetType(Reputation), "AddReputationForSoldGoods", MethodType.Normal)>
Public Class AddReputationPatch
    Public Shared Function Prefix(ByVal city As City, ByVal goodType As String, ByVal amountSold As Integer, ByVal shipCaptainMultiplier As Single) As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        ' Calculate reputation gain (doubled from 0.25f to 0.5f)
        Dim consumption = Mathf.Max(zip.lexy.tgame.simulation.consumption.Consumption.GetConsumption(city, goodType), 1.0F)
        Dim currentAmount = city.goods(goodType).amount
        Dim reputationGain = Mathf.Min(Mathf.Max(2.0F * consumption - currentAmount, 0.0F), CSng(amountSold)) * 0.5F * shipCaptainMultiplier

        ' Add reputation
        city.playerReputation += reputationGain

        ' Add mayor respect (25% of reputation gain)
        city.mayor.respect += Mathf.FloorToInt(reputationGain * 0.25F)

        Return False ' Skip original method
    End Function
End Class

' ============================================================
' DWELLING PATCHES - Make dwellings (22/23/24) player-buildable
' ============================================================

' Patch BuildingTypes.ForGood to handle dwelling resource names
<HarmonyPatch(GetType(BuildingTypes), "ForGood", MethodType.Normal)>
Public Class ForGoodPatch
    Public Shared Function Prefix(ByVal good As String, ByRef __result As Integer) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Select Case good
            Case "dwelling-low"
                __result = 22
                Return False
            Case "dwelling-mid"
                __result = 23
                Return False
            Case "dwelling-high"
                __result = 24
                Return False
            Case "shop-bakery"
                __result = 29
                Return False
            Case "shop-tavern"
                __result = 30
                Return False
            Case "shop-butcher"
                __result = 31
                Return False
            Case "shop-blacksmith"
                __result = 32
                Return False
            Case "shop-cooks-haven"
                __result = 33
                Return False
            Case "shop-tailor"
                __result = 34
                Return False
            Case "shop-carpenter"
                __result = 35
                Return False
            Case Else
                Return True
        End Select
    End Function
End Class

' ShowBusinessesPatch removed - dwellings and shops are now shown via FilterAllPatch,
' and Bonus filter shows only the city's bonus businesses without dwellings/shops.

' Patch BuildableBusiness.Initialize to handle dwelling and shop entries in the build menu
<HarmonyPatch(GetType(BuildableBusiness), "Initialize", MethodType.Normal)>
Public Class BuildableBusinessInitPatch
    Private Shared ReadOnly DwellingNames As New Dictionary(Of String, String) From {
        {"dwelling-low", "Low Dwelling"},
        {"dwelling-mid", "Mid Dwelling"},
        {"dwelling-high", "Upper Dwelling"}
    }

    Private Shared ReadOnly ShopTypes As New Dictionary(Of String, Integer) From {
        {"shop-bakery", 29},
        {"shop-tavern", 30},
        {"shop-butcher", 31},
        {"shop-blacksmith", 32},
        {"shop-cooks-haven", 33},
        {"shop-tailor", 34},
        {"shop-carpenter", 35}
    }

    Public Shared Function Prefix(ByVal __instance As BuildableBusiness, ByVal goodName As String) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance

        If DwellingNames.ContainsKey(goodName) Then
            ' Set the goodName field so OnPointerDown dispatches the correct resource
            GetType(BuildableBusiness).GetField("goodName", flags).SetValue(__instance, goodName)

            ' Set business name text
            Dim nameObj = GetType(BuildableBusiness).GetField("businessName", flags).GetValue(__instance)
            nameObj.GetType().GetProperty("text").SetValue(nameObj, DwellingNames(goodName))

            ' Set description text
            Dim descObj = GetType(BuildableBusiness).GetField("businessDescription", flags).GetValue(__instance)
            descObj.GetType().GetProperty("text").SetValue(descObj, "Residential building")

            ' Show representative icon for dwelling type
            Dim iconsObj = GetType(BuildableBusiness).GetField("goodsIcons", flags).GetValue(__instance)
            DirectCast(iconsObj, Component).gameObject.SetActive(True)
            Dim dwellingIcon = If(goodName = "dwelling-low", "logs", If(goodName = "dwelling-mid", "bricks", "stone"))
            DirectCast(iconsObj, GoodsIcons).ShowIcon(dwellingIcon)

            Return False
        End If

        If ShopTypes.ContainsKey(goodName) Then
            Dim shopType = ShopTypes(goodName)
            Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()
            Dim details = atlas.Get(shopType)

            ' Set the goodName field so OnPointerDown dispatches the correct resource
            GetType(BuildableBusiness).GetField("goodName", flags).SetValue(__instance, goodName)

            ' Set business name from atlas
            Dim nameObj = GetType(BuildableBusiness).GetField("businessName", flags).GetValue(__instance)
            nameObj.GetType().GetProperty("text").SetValue(nameObj, details.GetBuildingName())

            ' Set description from atlas
            Dim descObj = GetType(BuildableBusiness).GetField("businessDescription", flags).GetValue(__instance)
            descObj.GetType().GetProperty("text").SetValue(descObj, details.GetDescription())

            ' Show primary ingredient icon for shop
            Dim iconsObj = GetType(BuildableBusiness).GetField("goodsIcons", flags).GetValue(__instance)
            DirectCast(iconsObj, Component).gameObject.SetActive(True)
            If details.monthlyGoodCost IsNot Nothing AndAlso details.monthlyGoodCost.Count > 0 Then
                DirectCast(iconsObj, GoodsIcons).ShowIcon(details.monthlyGoodCost(0).goodType)
            End If

            Return False
        End If

        Return True
    End Function
End Class

' Patch SetShadowModel to clone an existing same-sized shadow for shops.
' Dwellings (22/23/24) have native shadows; shops (29-35) do not.
' Existing shadows have their mesh origin at the grid corner, which is required
' for correct alignment and in-place rotation.
' Size mapping: type 29 (2x3) -> clone "3" (Brickworks 2x3),
'               type 30 (3x3) -> clone "1" (Brewery 3x3),
'               type 31 (2x2) -> clone "4" (Clay Pit 2x2),
'               type 32 (2x2) -> clone "4" (Clay Pit 2x2),
'               type 33 (2x2) -> clone "4" (Clay Pit 2x2),
'               type 34 (2x3) -> clone "3" (Brickworks 2x3),
'               type 35 (2x3) -> clone "3" (Brickworks 2x3).
<HarmonyPatch(GetType(BuildingPlacementManager), "SetShadowModel", MethodType.Normal)>
Public Class SetShadowModelPatch
    Private Shared ReadOnly ShadowSourceMap As New Dictionary(Of Integer, String) From {
        {29, "3"},
        {30, "1"},
        {31, "4"},
        {32, "4"},
        {33, "4"},
        {34, "3"},
        {35, "3"}
    }

    Public Shared Sub Prefix(ByVal __instance As BuildingPlacementManager)
        If Not HSHPBuildMode.IsEnabled() Then Return
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance

        Dim buildingType = CInt(GetType(BuildingPlacementManager).GetField("placedBuildingType", flags).GetValue(__instance))
        If Not ShadowSourceMap.ContainsKey(buildingType) Then Return

        Dim buildingShadow = DirectCast(GetType(BuildingPlacementManager).GetField("buildingShadow", flags).GetValue(__instance), Transform)

        Dim childName = buildingType.ToString()
        If buildingShadow.Find(childName) IsNot Nothing Then Return

        Dim sourceId = ShadowSourceMap(buildingType)
        Dim sourceChild = buildingShadow.Find(sourceId)
        If sourceChild Is Nothing Then sourceChild = buildingShadow.Find("default")
        If sourceChild Is Nothing Then Return

        Dim clone = UnityEngine.Object.Instantiate(sourceChild.gameObject, buildingShadow)
        clone.name = childName
        clone.SetActive(False)
        clone.transform.localPosition = sourceChild.localPosition
        clone.transform.localRotation = sourceChild.localRotation
        clone.transform.localScale = sourceChild.localScale
    End Sub
End Class

' Patch GameState.OnPlayerBuildingCreationRequested to make dwellings town-owned.
' Must run BEFORE the original method because Initialize() caches isHuman from ownerType.
' Flow: PlaceBuilding -> dispatch event -> [OUR PREFIX changes ownership] -> 
'       GameState adds to city -> dispatches OnPlayerBuildingCreated ->
'       BuildingDisplayManager.SpawnBuilding -> Building.Initialize(details)
'       which reads details.ownerType to set isHuman (determines click behavior).
<HarmonyPatch(GetType(GameState), "OnPlayerBuildingCreationRequested", MethodType.Normal)>
Public Class DwellingOwnershipPatch
    Public Shared Sub Prefix(ByVal __instance As GameState, ByVal evt As Object)
        If Not HSHPBuildMode.IsEnabled() Then Return
        ' Access evt.building via reflection to avoid Building type ambiguity
        Dim buildingField = evt.GetType().GetField("building")
        If buildingField Is Nothing Then Return

        Dim building = buildingField.GetValue(evt)
        If building Is Nothing Then Return

        Dim typeField = building.GetType().GetField("type")
        Dim buildingType = CInt(typeField.GetValue(building))

        If buildingType <> 22 AndAlso buildingType <> 23 AndAlso buildingType <> 24 Then Return

        ' Set ownership to city (OwnerType.CITY = 0)
        Dim ownerTypeField = building.GetType().GetField("ownerType")
        Dim ownerIdField = building.GetType().GetField("ownerId")

        ownerTypeField.SetValue(building, 0)

        ' ownerId = city index in the cities list (matches how game creates town dwellings)
        Dim cityIndex = __instance.cities.IndexOf(__instance.viewCity)
        ownerIdField.SetValue(building, cityIndex)
    End Sub
End Class

' Patch GameState.OnPlayerBuildingCreationRequested to grant reputation and mayor respect
' when any dwelling (22/23/24) is placed. +100 reputation, +50 mayor respect.
<HarmonyPatch(GetType(GameState), "OnPlayerBuildingCreationRequested", MethodType.Normal)>
Public Class DwellingPlacementRewardPatch
    Public Shared Sub Postfix(ByVal __instance As GameState, ByVal evt As Object)
        If Not HSHPDifficulty.IsHardMode() Then Return
        Dim buildingField = evt.GetType().GetField("building")
        If buildingField Is Nothing Then Return

        Dim building = buildingField.GetValue(evt)
        If building Is Nothing Then Return

        Dim buildingType = CInt(building.GetType().GetField("type").GetValue(building))
        If buildingType <> 22 AndAlso buildingType <> 23 AndAlso buildingType <> 24 Then Return

        ' Grant 100 reputation and 50 mayor respect
        __instance.viewCity.playerReputation += 100.0F
        __instance.viewCity.mayor.respect += 50
    End Sub
End Class

' Patch SetStatusAfterCompletion to set dwellings to OPERATIONAL instead of MISSING_RESOURCES
' and all shops to MISSING_RESOURCES (2) instead of AUCTION (3) to skip the auction system.
<HarmonyPatch(GetType(BuildingTurnAdvanceManager), "SetStatusAfterCompletion", MethodType.Normal)>
Public Class DwellingStatusPatch
    Public Shared Function Prefix(ByVal building As Object) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Dim buildingType = CInt(building.GetType().GetField("type").GetValue(building))

        ' Dwellings (22/23/24) should be OPERATIONAL (0) after construction
        If buildingType = 22 OrElse buildingType = 23 OrElse buildingType = 24 Then
            building.GetType().GetField("status").SetValue(building, 0)
            Return False
        End If

        ' All shops skip auction, go straight to MISSING_RESOURCES (2)
        If BuildingTypes.IsShop(buildingType) Then
            building.GetType().GetField("status").SetValue(building, 2)
            Return False
        End If

        Return True
    End Function
End Class

' ============================================================
' SHOP PATCHES - Make shops (29-35) player-buildable
' ============================================================

' Patch GameState.OnPlayerBuildingCreationRequested to create a Shop object
' for player-built shops. Sets managerId to human and calculates monthlyIncome
' using the auction bid formula (1.6 x sum of ingredient costs).
<HarmonyPatch(GetType(GameState), "OnPlayerBuildingCreationRequested", MethodType.Normal)>
Public Class ShopCreationPatch
    Private Shared ReadOnly ShopTypeIds As New HashSet(Of Integer) From {29, 30, 31, 32, 33, 34, 35}

    Public Shared Sub Postfix(ByVal __instance As GameState, ByVal evt As Object)
        If Not HSHPBuildMode.IsEnabled() Then Return
        Dim buildingField = evt.GetType().GetField("building")
        If buildingField Is Nothing Then Return

        Dim building = buildingField.GetValue(evt)
        If building Is Nothing Then Return

        Dim buildingType = CInt(building.GetType().GetField("type").GetValue(building))
        If Not ShopTypeIds.Contains(buildingType) Then Return

        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()
        Dim details = atlas.Get(buildingType)

        ' Calculate monthlyIncome using auction formula: 1.6 x sum(monthlyGoodCost x corePrices)
        Dim costSum As Single = 0.0F
        For Each gc As GoodQuantity In details.monthlyGoodCost
            costSum += CSng(gc.quantity) * __instance.corePrices(gc.goodType)
        Next
        Dim monthlyIncome = Mathf.RoundToInt(1.6F * costSum)

        ' Create Shop object and add to city
        Dim shop As New Shop()
        shop.id = __instance.GetUniqueId()
        shop.building = DirectCast(building, zip.lexy.tgame.state.building.Building)
        shop.managerId = __instance.human.id
        shop.monthlyIncome = monthlyIncome

        __instance.viewCity.shops.Add(shop)
    End Sub
End Class

' Patch SetWindowDetails to handle dwellings without crashing on null producedItem.
' Hides goods icon, monthly costs, and production sections. Shows build costs normally.
<HarmonyPatch(GetType(BuildWindowDetails), "SetWindowDetails", MethodType.Normal)>
Public Class DwellingWindowDetailsPatch
    Private Shared ReadOnly DwellingTypes As New HashSet(Of Integer) From {22, 23, 24}

    Public Shared Function Prefix(ByVal __instance As BuildWindowDetails, ByVal buildingDetails As BuildingScriptableObject) As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim bwdType = GetType(BuildWindowDetails)

        ' Always re-enable goodsIcons so normal buildings show their icon
        ' (dwellings disable it, which persists across detail views)
        Dim goodsIcons = DirectCast(bwdType.GetField("goodsIcons", flags).GetValue(__instance), Component)
        goodsIcons.gameObject.SetActive(True)

        ' Also re-enable monthly cost fields that dwellings hide
        DirectCast(bwdType.GetField("monthlyCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(True)
        DirectCast(bwdType.GetField("monthlyWinterCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(True)
        DirectCast(bwdType.GetField("monthlyCostItem1", flags).GetValue(__instance), Component).gameObject.SetActive(True)
        DirectCast(bwdType.GetField("monthlyCostItem2", flags).GetValue(__instance), Component).gameObject.SetActive(True)
        DirectCast(bwdType.GetField("monthlyCostItem3", flags).GetValue(__instance), Component).gameObject.SetActive(True)

        ' Check if this is a dwelling by looking up its key in the atlas
        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()
        Dim key = atlas.GetKey(buildingDetails)

        Dim isDwelling = False
        For Each dt In DwellingTypes
            If key = dt.ToString() Then
                isDwelling = True
                Exit For
            End If
        Next

        ' Check if this is a shop
        Dim shopTypeId As Integer = 0
        Dim isShop = Integer.TryParse(key, shopTypeId) AndAlso BuildingTypes.IsShop(shopTypeId)

        If Not isDwelling AndAlso Not isShop Then Return True

        If isShop Then
            ' ---- Shop custom handler ----
            Dim shopTitle = DirectCast(bwdType.GetField("title", flags).GetValue(__instance), Object)
            shopTitle.GetType().GetProperty("text").SetValue(shopTitle, buildingDetails.GetBuildingName())

            Dim shopDesc = DirectCast(bwdType.GetField("description", flags).GetValue(__instance), Object)
            shopDesc.GetType().GetProperty("text").SetValue(shopDesc, buildingDetails.GetDescription())

            ' Show primary ingredient icon for shop
            If buildingDetails.monthlyGoodCost IsNot Nothing AndAlso buildingDetails.monthlyGoodCost.Count > 0 Then
                DirectCast(goodsIcons, GoodsIcons).ShowIcon(buildingDetails.monthlyGoodCost(0).goodType)
            Else
                goodsIcons.gameObject.SetActive(False)
            End If

            ' Hide employment (shops don't employ pops in game logic)
            DirectCast(bwdType.GetField("lowerClass", flags).GetValue(__instance), Component).gameObject.SetActive(False)
            DirectCast(bwdType.GetField("middleClass", flags).GetValue(__instance), Component).gameObject.SetActive(False)
            DirectCast(bwdType.GetField("upperClass", flags).GetValue(__instance), Component).gameObject.SetActive(False)
            DirectCast(bwdType.GetField("missingWorkersWarning", flags).GetValue(__instance), Component).gameObject.SetActive(False)

            ' Construction time
            Dim ctObj = DirectCast(bwdType.GetField("constructionTime", flags).GetValue(__instance), Object)
            Dim ctText = Localization.ForKey("completion-time", Nothing)
            ctText = ctText.Replace("{weekCount}", (CDbl(buildingDetails.completionTimeInTurns) / 2.0).ToString())
            ctText = ctText.Replace("{turnCount}", buildingDetails.completionTimeInTurns.ToString())
            ctObj.GetType().GetProperty("text").SetValue(ctObj, ctText)

            ' Build costs
            Dim shopTradingMgr = InstanceProvider.GetInstance(Of TradingSimulationManager)()
            Dim getReqAmt = bwdType.GetMethod("GetRequiredAmount", flags)
            Dim bldGoods = DirectCast(GetType(BuildWindowDetails).GetField("BUILDING_GOODS", BindingFlags.Public Or BindingFlags.Static).GetValue(Nothing), List(Of String))
            Dim reqFunc As Func(Of String, Integer) = Function(g As String) As Integer
                                                          Return CInt(getReqAmt.Invoke(__instance, New Object() {g, buildingDetails}))
                                                      End Function
            Dim shopTransactions = shopTradingMgr.GetAllTransactionsForGoods(bldGoods, reqFunc, False)
            bwdType.GetMethod("SetBuildCosts", flags).Invoke(__instance, New Object() {buildingDetails, shopTransactions})

            ' Calculate monthly income (same formula as ShopCreationPatch: 1.6 x ingredient costs)
            Dim gameState = InstanceProvider.GetInstance(Of GameState)()
            Dim costSum As Single = 0.0F
            For Each gc As GoodQuantity In buildingDetails.monthlyGoodCost
                costSum += CSng(gc.quantity) * gameState.corePrices(gc.goodType)
            Next
            Dim monthlyIncome = Mathf.RoundToInt(1.6F * costSum)

            ' Hide monthly cost money (income already shown in production section)
            DirectCast(bwdType.GetField("monthlyCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(False)

            ' Monthly cost items: show ingredients
            Dim setMCI = bwdType.GetMethod("SetMonthlyCostItem", flags)
            setMCI.Invoke(__instance, New Object() {buildingDetails, 0, bwdType.GetField("monthlyCostItem1", flags).GetValue(__instance)})
            setMCI.Invoke(__instance, New Object() {buildingDetails, 1, bwdType.GetField("monthlyCostItem2", flags).GetValue(__instance)})
            setMCI.Invoke(__instance, New Object() {buildingDetails, 2, bwdType.GetField("monthlyCostItem3", flags).GetValue(__instance)})

            ' Hide winter cost (shops not seasonal)
            DirectCast(bwdType.GetField("monthlyWinterCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(False)

            ' Production: show monthly income matching existing shop display
            Dim prod1 = bwdType.GetField("production1", flags).GetValue(__instance)
            DirectCast(prod1, Component).gameObject.SetActive(True)
            Dim incomeText = Localization.ForKey("monthy-coin-cost", Nothing).Replace("{coins}", monthlyIncome.ToString())
            prod1.GetType().GetMethod("UpdateItem").Invoke(prod1, New Object() {"coin", 0.0F, incomeText, ""})
            prod1.GetType().GetMethod("SwitchUOM").Invoke(prod1, New Object() {"coin"})

            ' Hide other production fields
            DirectCast(bwdType.GetField("productionWinter", flags).GetValue(__instance), Component).gameObject.SetActive(False)
            DirectCast(bwdType.GetField("bonusProductionDetails", flags).GetValue(__instance), Component).gameObject.SetActive(False)
            DirectCast(bwdType.GetField("winterProductionWarning", flags).GetValue(__instance), Component).gameObject.SetActive(False)

            ' Setup build button
            bwdType.GetMethod("SetupBuildButton", flags).Invoke(__instance, New Object() {buildingDetails, shopTransactions})

            Return False
        End If

        ' Title and description
        Dim title = DirectCast(bwdType.GetField("title", flags).GetValue(__instance), Object)
        title.GetType().GetProperty("text").SetValue(title, buildingDetails.GetBuildingName())

        Dim desc = DirectCast(bwdType.GetField("description", flags).GetValue(__instance), Object)
        desc.GetType().GetProperty("text").SetValue(desc, buildingDetails.GetDescription())

        ' Show representative icon for dwelling type
        Dim dwellingDetailIcon = If(key = "22", "logs", If(key = "23", "bricks", "stone"))
        DirectCast(goodsIcons, GoodsIcons).ShowIcon(dwellingDetailIcon)

        ' Set employed pops
        Dim setEmployedPops = bwdType.GetMethod("SetEmployedPops", flags)
        setEmployedPops.Invoke(__instance, New Object() {buildingDetails})

        ' Set workers warning
        Dim setWorkersWarning = bwdType.GetMethod("SetWorkersWarning", flags)
        setWorkersWarning.Invoke(__instance, New Object() {buildingDetails})

        ' Construction time
        Dim constructionTime = DirectCast(bwdType.GetField("constructionTime", flags).GetValue(__instance), Object)
        Dim timeText = Localization.ForKey("completion-time", Nothing)
        timeText = timeText.Replace("{weekCount}", (CDbl(buildingDetails.completionTimeInTurns) / 2.0).ToString())
        timeText = timeText.Replace("{turnCount}", buildingDetails.completionTimeInTurns.ToString())
        constructionTime.GetType().GetProperty("text").SetValue(constructionTime, timeText)

        ' Build costs (uses trading simulation to calculate material costs)
        Dim tradingMgr = InstanceProvider.GetInstance(Of TradingSimulationManager)()
        Dim getRequiredAmount = bwdType.GetMethod("GetRequiredAmount", flags)
        Dim buildingGoods = GetType(BuildWindowDetails).GetField("BUILDING_GOODS", BindingFlags.Public Or BindingFlags.Static).GetValue(Nothing)
        Dim goodsList = DirectCast(buildingGoods, List(Of String))

        ' Build the Func(Of String, Integer) delegate for GetAllTransactionsForGoods
        Dim reqAmountFunc As Func(Of String, Integer) = Function(good As String) As Integer
                                                            Return CInt(getRequiredAmount.Invoke(__instance, New Object() {good, buildingDetails}))
                                                        End Function

        Dim transactions = tradingMgr.GetAllTransactionsForGoods(goodsList, reqAmountFunc, False)

        Dim setBuildCosts = bwdType.GetMethod("SetBuildCosts", flags)
        setBuildCosts.Invoke(__instance, New Object() {buildingDetails, transactions})

        ' Hide monthly costs (dwellings have no monthly costs)
        DirectCast(bwdType.GetField("monthlyCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(False)
        DirectCast(bwdType.GetField("monthlyWinterCostMoney", flags).GetValue(__instance), Component).gameObject.SetActive(False)
        DirectCast(bwdType.GetField("monthlyCostItem1", flags).GetValue(__instance), Component).gameObject.SetActive(False)
        DirectCast(bwdType.GetField("monthlyCostItem2", flags).GetValue(__instance), Component).gameObject.SetActive(False)
        DirectCast(bwdType.GetField("monthlyCostItem3", flags).GetValue(__instance), Component).gameObject.SetActive(False)

        ' Hide production fields (dwellings don't produce goods)
        Dim hideProduction = bwdType.GetMethod("HideProductionFields", flags)
        hideProduction.Invoke(__instance, Nothing)

        ' Setup build button (bypass building slot limit for dwellings)
        DwellingBuildBypass.IsBuildingDwelling = True
        Dim setupBuildButton = bwdType.GetMethod("SetupBuildButton", flags)
        setupBuildButton.Invoke(__instance, New Object() {buildingDetails, transactions})
        DwellingBuildBypass.IsBuildingDwelling = False

        Return False
    End Function
End Class

' Static flag to bypass building slot check when viewing dwellings
Public Module DwellingBuildBypass
    Public IsBuildingDwelling As Boolean = False
End Module

' Patch DoesPlayerHaveEnoughReputationToBuild to always allow dwelling construction
' and exclude dwellings (ownerType=0) from the player building count.
<HarmonyPatch(GetType(BuildWindowDetails), "DoesPlayerHaveEnoughReputationToBuild", MethodType.Normal)>
Public Class DwellingSlotBypassPatch
    Public Shared Function Prefix(ByRef __result As Boolean) As Boolean
        If DwellingBuildBypass.IsBuildingDwelling Then
            __result = True
            Return False
        End If
        Return True
    End Function
End Class

' Shared blocked-goals set used by both patches
Public Module BlockedGoals
    Public ReadOnly Blocked As New HashSet(Of String) From {
        "build_house", "build_shop"
    }

    Public ReadOnly Replacements As String() = {
        "give_bonus",
        "give_shipyard_bonus"
    }

    Private _rng As New System.Random()

    Public Function RandomReplacement() As String
        Return Replacements(_rng.Next(Replacements.Length))
    End Function
End Module

' Patch Mayor.GetGoalTypes to remove blocked goals from initial population
<HarmonyPatch(GetType(Mayor), "GetGoalTypes", MethodType.Normal)>
Public Class RemoveBuildGoalsPatch
    Public Shared Sub Postfix(ByRef __result As List(Of String))
        If Not HSHPBuildMode.IsEnabled() Then Return
        For i = 0 To __result.Count - 1
            If BlockedGoals.Blocked.Contains(__result(i)) Then
                __result(i) = BlockedGoals.RandomReplacement()
            End If
        Next
    End Sub
End Class

' Patch GoalGiver.GiveGoal to intercept blocked goal types during refresh/replacement
' This covers MayorGoalSelectionWindow.SetNewGoals which uses ALL_EXCEPT_SHIPYARD_UPGRADE_AND_CITY_GROW
' and does NOT go through Mayor.GetGoalTypes
<HarmonyPatch(GetType(GoalGiver), "GiveGoal", MethodType.Normal)>
Public Class GiveGoalPatch
    Public Shared Sub Prefix(ByRef goalType As String)
        If Not HSHPBuildMode.IsEnabled() Then Return
        If BlockedGoals.Blocked.Contains(goalType) Then
            goalType = BlockedGoals.RandomReplacement()
        End If
    End Sub
End Class

' ============================================================
' POPULATION CAP - Limit growth to available dwelling space
' ============================================================

' Patch AdjustPopCountBasedOnHappiness to cap population at dwelling capacity.
' Population can still decline naturally but cannot grow beyond available housing.
' Dwelling capacity = count of operational (status=0) dwellings of that type x pops per dwelling.
<HarmonyPatch(GetType(PopTurnAdvanceManager), "AdjustPopCountBasedOnHappiness", MethodType.Normal)>
Public Class PopulationCapPatch
    Public Shared Sub Postfix(ByVal __instance As Object, ByVal city As City)
        If Not HSHPPersonal.IsEnabled() Then Return
        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()

        ' Calculate capacity for each pop class from operational dwellings
        Dim lowerCap = city.buildings.Where(Function(b) b.type = 22 AndAlso b.status = 0).Count() * atlas.Get(22).requiredLowerPops
        Dim middleCap = city.buildings.Where(Function(b) b.type = 23 AndAlso b.status = 0).Count() * atlas.Get(23).requiredMiddlePops
        Dim upperCap = city.buildings.Where(Function(b) b.type = 24 AndAlso b.status = 0).Count() * atlas.Get(24).requiredUpperPops

        ' Cap population at dwelling capacity (only prevents growth, decline still allowed)
        If city.lowerPop > lowerCap Then city.lowerPop = lowerCap
        If city.middlePop > middleCap Then city.middlePop = middleCap
        If city.higherPop > upperCap Then city.higherPop = upperCap

        ' Ensure min-pop floors don't exceed capacity either
        city.minLowPop = Math.Min(city.minLowPop, lowerCap)
        city.minMidPop = Math.Min(city.minMidPop, middleCap)
        city.minHighPop = Math.Min(city.minHighPop, upperCap)
    End Sub
End Class

' ============================================================
' HOUSING OCCUPANCY - Show dwelling capacity in Trader's Hall
' ============================================================

' Patch TradersHallOverviewWindow.Show to append housing occupancy info
' below the employment section, using the same format: pop/capacity (xx%)
' Also pushes the traders section down to avoid overlap.
<HarmonyPatch(GetType(TradersHallOverviewWindow), "Show", MethodType.Normal)>
Public Class HousingOccupancyPatch
    Public Shared Sub Postfix(ByVal __instance As TradersHallOverviewWindow)
        If Not HSHPPersonal.IsEnabled() Then Return
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim wndType = GetType(TradersHallOverviewWindow)

        Dim employmentField = wndType.GetField("employment", flags)
        Dim employmentText = employmentField.GetValue(__instance)
        Dim textProp = employmentText.GetType().GetProperty("text")
        Dim currentText = CStr(textProp.GetValue(employmentText))

        Dim city = InstanceProvider.GetInstance(Of GameState)().viewCity
        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()

        Dim lowerCap = city.buildings.Where(Function(b) b.type = 22 AndAlso b.status = 0).Count() * atlas.Get(22).requiredLowerPops
        Dim middleCap = city.buildings.Where(Function(b) b.type = 23 AndAlso b.status = 0).Count() * atlas.Get(23).requiredMiddlePops
        Dim upperCap = city.buildings.Where(Function(b) b.type = 24 AndAlso b.status = 0).Count() * atlas.Get(24).requiredUpperPops

        Dim lowerPct = If(lowerCap > 0, CInt(CSng(city.lowerPop) / CSng(lowerCap) * 100.0F), 0)
        Dim middlePct = If(middleCap > 0, CInt(CSng(city.middlePop) / CSng(middleCap) * 100.0F), 0)
        Dim upperPct = If(upperCap > 0, CInt(CSng(city.higherPop) / CSng(upperCap) * 100.0F), 0)

        Dim housingText = String.Concat(
            vbLf, vbLf,
            "Housing Occupancy:", vbLf, vbLf,
            Localization.ForKey("lower-class", Nothing), ": ",
            city.lowerPop.ToString(), "/", lowerCap.ToString(),
            " (", lowerPct.ToString(), "% occupied)", vbLf,
            Localization.ForKey("middle-class", Nothing), ": ",
            city.middlePop.ToString(), "/", middleCap.ToString(),
            " (", middlePct.ToString(), "% occupied)", vbLf,
            Localization.ForKey("upper-class", Nothing), ": ",
            city.higherPop.ToString(), "/", upperCap.ToString(),
            " (", upperPct.ToString(), "% occupied)"
        )

        textProp.SetValue(employmentText, currentText & housingText)

        ' Hide the traders list text
        Dim tradersField = wndType.GetField("traders", flags)
        Dim tradersObj = tradersField.GetValue(__instance)
        DirectCast(tradersObj, Component).gameObject.SetActive(False)

        ' Hide the traders section title by searching sibling TMP_Text components
        ' whose text content contains "trader" (case-insensitive)
        Dim tradersTransform = DirectCast(tradersObj, Component).transform
        If tradersTransform.parent IsNot Nothing Then
            For i = 0 To tradersTransform.parent.childCount - 1
                Dim sibling = tradersTransform.parent.GetChild(i)
                If sibling Is tradersTransform Then Continue For
                Dim tmpTexts = sibling.GetComponentsInChildren(Of TMPro.TMP_Text)(True)
                For Each tmp In tmpTexts
                    If tmp.text IsNot Nothing AndAlso tmp.text.ToLower().Contains("trader") Then
                        sibling.gameObject.SetActive(False)
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub
End Class

' ============================================================
' HOUSING HAPPINESS - Occupancy-based happiness effect
' ============================================================

' Patch PopHappiness.SetupHomes to use occupancy rate instead of homelessness.
' Higher occupancy = greater negative effect on happiness.
' occupancy = totalPop / dwellingCapacity (clamped to 0-1)
' At low occupancy (<=50%): small positive bonus (+0.25)
' At high occupancy (100%): strong negative malus (uses existing homeEffectRange.x)
<HarmonyPatch(GetType(PopHappiness), "SetupHomes", MethodType.Normal)>
Public Class OccupancyHappinessPatch
    Public Shared Function Prefix(ByVal pops As Population, ByVal city As City, ByVal happiness As PopHappiness) As Boolean
        Dim isPersonal = HSHPPersonal.IsEnabled()
        Dim isBD = HSHPDifficulty.IsHardMode()
        If Not isPersonal AndAlso Not isBD Then Return True

        Dim hcProp = GetType(PopHappiness).GetProperty("homeCoverage", BindingFlags.Public Or BindingFlags.Instance)
        Dim heProp = GetType(PopHappiness).GetProperty("homesEffect", BindingFlags.Public Or BindingFlags.Instance)

        ' Zero-pop fix (for both FB's Changes and BD modes)
        Dim popCount = pops.GetPopsForCity(city)
        If popCount <= 0 Then
            hcProp.SetValue(happiness, 0.0F)
            heProp.SetValue(happiness, 0.0F)
            Return False
        End If

        ' Custom occupancy logic only applies when FB's Changes is enabled
        If Not isPersonal Then Return True

        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()
        Dim homeType = pops.GetHomeType()

        ' Calculate dwelling capacity
        Dim capacity As Integer
        If homeType = 22 Then
            capacity = city.buildings.Where(Function(b) b.type = 22 AndAlso b.status = 0).Count() * atlas.Get(22).requiredLowerPops
        ElseIf homeType = 23 Then
            capacity = city.buildings.Where(Function(b) b.type = 23 AndAlso b.status = 0).Count() * atlas.Get(23).requiredMiddlePops
        Else
            capacity = city.buildings.Where(Function(b) b.type = 24 AndAlso b.status = 0).Count() * atlas.Get(24).requiredUpperPops
        End If

        Dim occupancy As Single
        If capacity > 0 Then
            occupancy = Mathf.Clamp01(CSng(popCount) / CSng(capacity))
        Else
            occupancy = 1.0F
        End If

        ' Store occupancy as homeCoverage (repurposed: 1 = full occupancy = bad)
        ' homeCoverage is used by PopulationHappinessWindow.SetupHomeless for display
        hcProp.SetValue(happiness, occupancy)

        ' Calculate effect: high occupancy = negative, low occupancy = positive
        ' Lerp from positive (at 0% occupancy) to negative (at 100% occupancy)
        ' Custom malus values: Lower=-1.5 (was -1.8), Middle=-1.0, Upper=-1.2 (was -1.5)
        Dim effectRange As Vector2
        If homeType = 22 Then
            effectRange = New Vector2(-1.5F, 0.25F)
        ElseIf homeType = 23 Then
            effectRange = New Vector2(-1.0F, 0.25F)
        Else
            effectRange = New Vector2(-1.2F, 0.35F)
        End If
        Dim effect = Mathf.Lerp(effectRange.y, effectRange.x, occupancy)

        heProp.SetValue(happiness, effect)

        Return False
    End Function
End Class

' Patch PopulationHappinessWindow.SetupHomeless to show occupancy text
' instead of "x% homeless". Shows "x% housing occupancy" with effect value.
<HarmonyPatch(GetType(PopulationHappinessWindow), "SetupHomeless", MethodType.Normal)>
Public Class OccupancyDisplayPatch
    Public Shared Function Prefix(ByVal __instance As PopulationHappinessWindow, ByVal happiness As PopHappiness) As Boolean
        If Not HSHPPersonal.IsEnabled() Then Return True
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim wndType = GetType(PopulationHappinessWindow)

        ' homeCoverage is now repurposed as occupancy rate (0-1)
        Dim occupancyPct = CInt(happiness.homeCoverage * 100.0F)

        Dim labelField = wndType.GetField("homelessLabel", flags)
        Dim labelObj = labelField.GetValue(__instance)
        labelObj.GetType().GetProperty("text").SetValue(labelObj,
            occupancyPct.ToString() & "% housing occupancy")

        Dim effectField = wndType.GetField("homelessEffect", flags)
        Dim effectObj = effectField.GetValue(__instance)
        Dim effectText = If(happiness.homesEffect > 0.0F, "+", "") &
            Math.Round(CDbl(happiness.homesEffect), 2).ToString()
        effectObj.GetType().GetProperty("text").SetValue(effectObj, effectText)

        Return False
    End Function
End Class

' ============================================================
' ZERO-POP HAPPINESS FIXES - Prevent NaN at zero population
' ============================================================

' Patch SetupShops to return 0 effect when totalPop is 0.
' Original: (operationalShops * 50) / totalPop -> NaN when totalPop=0.
<HarmonyPatch(GetType(PopHappiness), "SetupShops", MethodType.Normal)>
Public Class ShopHappinessPatch
    Public Shared Function Prefix(ByVal city As City, ByVal happiness As PopHappiness) As Boolean
        If Not HSHPPersonal.IsEnabled() AndAlso Not HSHPDifficulty.IsHardMode() Then Return True
        If city.totalPop < 1 Then
            Dim seProp = GetType(PopHappiness).GetProperty("shopEffect", BindingFlags.Public Or BindingFlags.Instance)
            seProp.SetValue(happiness, 0.0F)
            Return False
        End If
        Return True
    End Function
End Class

' Patch SetupEmployment to use custom effect ranges and handle zero population.
' Original: employedPops / popsForCity -> NaN when popsForCity=0.
' Custom malus: Lower=-1.5 (was -2.0), Middle=-1.5, Upper=-1.0 (unchanged)
<HarmonyPatch(GetType(PopHappiness), "SetupEmployment", MethodType.Normal)>
Public Class EmploymentHappinessPatch
    Public Shared Function Prefix(ByVal pops As Population, ByVal city As City, ByVal happiness As PopHappiness) As Boolean
        Dim isPersonal = HSHPPersonal.IsEnabled()
        Dim isBD = HSHPDifficulty.IsHardMode()
        If Not isPersonal AndAlso Not isBD Then Return True

        Dim erProp = GetType(PopHappiness).GetProperty("employmentRate", BindingFlags.Public Or BindingFlags.Instance)
        Dim eeProp = GetType(PopHappiness).GetProperty("employmentEffect", BindingFlags.Public Or BindingFlags.Instance)

        ' Zero-pop fix (for both FB's Changes and BD modes)
        If pops.GetPopsForCity(city) < 1 Then
            erProp.SetValue(happiness, 1.0F)
            eeProp.SetValue(happiness, 0.0F)
            Return False
        End If

        ' Custom effect ranges only apply when FB's Changes is enabled
        If Not isPersonal Then Return True

        ' Calculate employment rate (same as original)
        Dim employed = PopHappiness.GetEmployedPopsInCity(pops, city)
        Dim rate = Mathf.Min(CSng(employed) / CSng(pops.GetPopsForCity(city)), 1.0F)
        erProp.SetValue(happiness, rate)

        ' Custom effect ranges: Lower=(-1.5, 0.4), Middle=(-1.5, 0.4), Upper=(-1.0, 0.4)
        Dim homeType = pops.GetHomeType()
        Dim effectRange As Vector2
        If homeType = 22 Then
            effectRange = New Vector2(-1.5F, 0.4F)
        ElseIf homeType = 23 Then
            effectRange = New Vector2(-1.5F, 0.4F)
        Else
            effectRange = New Vector2(-1.0F, 0.4F)
        End If
        Dim effect = Mathf.Lerp(effectRange.x, effectRange.y, rate)
        eeProp.SetValue(happiness, effect)

        Return False
    End Function
End Class

' Patch GetEmploymentLabelForPops to return 0% when popCount is 0.
' Original: employedPops / popCount * 100 -> negative/NaN when popCount=0.
<HarmonyPatch(GetType(TradersHallOverviewWindow), "GetEmploymentLabelForPops", MethodType.Normal)>
Public Class EmploymentLabelZeroPopPatch
    Public Shared Function Prefix(ByVal city As City, ByVal population As Population, ByVal popCount As Integer, ByVal labelKey As String, ByRef __result As String) As Boolean
        If Not HSHPPersonal.IsEnabled() AndAlso Not HSHPDifficulty.IsHardMode() Then Return True
        If popCount < 1 Then
            __result = String.Concat(
                Localization.ForKey(labelKey, Nothing), ": ",
                "0/0 (",
                Localization.ForKey("employment-percent", Nothing).Replace("{percentage}", "0"),
                ")")
            Return False
        End If
        Return True
    End Function
End Class

' ============================================================
' SHOP PROTECTION - Prevent warnings, loss, and auctions
' ============================================================

' Patch DealWithShopMissingResources to prevent mayor warnings and shop loss.
' Without this, after 24 turns of missing ingredients the mayor takes the shop away,
' reduces reputation, and creates an auction.
<HarmonyPatch(GetType(BuildingTurnAdvanceManager), "DealWithShopMissingResources", MethodType.Normal)>
Public Class NoShopMissingResourcesPatch
    Public Shared Function Prefix() As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Return False
    End Function
End Class

' Patch LoseRobotShopIfNeeded to prevent random robot-managed shop auctions.
' Without this, robot traders randomly lose shops (0.3% per turn) triggering re-auctions.
<HarmonyPatch(GetType(BuildingTurnAdvanceManager), "LoseRobotShopIfNeeded", MethodType.Normal)>
Public Class NoRobotShopLossPatch
    Public Shared Function Prefix() As Boolean
        If Not HSHPBuildMode.IsEnabled() Then Return True
        Return False
    End Function
End Class

' ============================================================
' NEW GAME SETUP - Strip cities and remove AI traders
' ============================================================

' Patch SetInitialCities to customize city starting conditions.
' Runs after all cities have been initialized (businesses, dwellings, town buildings placed).
' - Removes all AI traders (keeps only the human player)
' - All cities: removes all businesses/dwellings, zeroes population
' - Home city: sets 250 stockpile + dwelling construction materials
' - Non-player cities: zeroes stockpiles
' - Sets starting funds to 50K
<HarmonyPatch(GetType(GameState), "SetInitialCities", MethodType.Normal)>
Public Class NewGameSetupPatch
    Private Shared ReadOnly AllGoods As String() = {
        "ale", "beef", "bricks", "clay", "cloth", "fish", "grain", "honey",
        "iron-bars", "logs", "lumber", "mead", "ore", "pottery", "salt",
        "stone", "tar", "vegetables", "wine", "wooden-tools", "wool"
    }

    Public Shared Sub Postfix(ByVal __instance As GameState)
        If Not HSHPDifficulty.IsHardMode() Then Return

        ' Remove all AI traders (keep only the human player)
        Dim humanId = __instance.human.id
        __instance.traders.RemoveAll(Function(t) t.id <> humanId)

        Dim playerIdx = __instance.playerCityIdx

        For i = 0 To __instance.cities.Count - 1
            Dim city = __instance.cities(i)

            ' Remove all production buildings (1-21) and dwellings (22-24)
            ' Keeps town buildings: TradersHall(25), TownHall(26), Park(27), Market(28)
            city.buildings.RemoveAll(Function(b) b.type < 25)
            city.shops.Clear()

            ' Re-place remaining town buildings on the grid now that production/dwellings are gone
            BuildingPlacement.PlaceOptimally(city)

            ' Zero population
            city.lowerPop = 0
            city.middlePop = 0
            city.higherPop = 0
            city.minLowPop = 0
            city.minMidPop = 0
            city.minHighPop = 0

            If i = playerIdx Then
                ' Home city: set 250 stockpile of every good
                For Each goodName In AllGoods
                    city.goods(goodName) = New ItemStack(goodName, 250.0F, 0.0F)
                Next

                ' Add construction materials for 3 of each dwelling type to town stockpile
                Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()
                For Each dwellingType In New Integer() {22, 23, 24}
                    Dim details = atlas.Get(dwellingType)
                    If details.buildCost IsNot Nothing Then
                        For Each gc As GoodQuantity In details.buildCost
                            Dim needed = CSng(gc.quantity) * 3.0F
                            If city.goods.ContainsKey(gc.goodType) Then
                                Dim existing = city.goods(gc.goodType).amount
                                city.goods(gc.goodType) = New ItemStack(gc.goodType, existing + needed, 0.0F)
                            Else
                                city.goods(gc.goodType) = New ItemStack(gc.goodType, needed, 0.0F)
                            End If
                        Next
                    End If
                Next
            Else
                ' Zero stockpiles for non-player cities
                For Each goodName In AllGoods
                    If city.goods.ContainsKey(goodName) Then
                        city.goods(goodName) = New ItemStack(goodName, 0.0F, 0.0F)
                    End If
                Next
            End If
        Next

        ' Set starting funds to 50K
        __instance.human.balance = 50000L
    End Sub
End Class

' Patch City.Initialize to replace city bonuses with exactly 1 raw good (no prerequisites).
' Instead of filtering the existing 3 bonuses (which can yield 0 if all are processed goods),
' picks deterministically from all raw goods based on city id.
' Only active when FB's Changes is enabled.
<HarmonyPatch(GetType(City), "Initialize", MethodType.Normal)>
Public Class CityBonusFilterPatch
    Private Shared ReadOnly AllGoods As List(Of String) = New List(Of String) From {
        "ale", "beef", "bricks", "clay", "cloth", "fish", "grain", "honey",
        "iron-bars", "logs", "lumber", "mead", "ore", "pottery", "salt",
        "stone", "tar", "vegetables", "wine", "wooden-tools", "wool"}

    Public Shared Sub Prefix(ByVal id As Integer, ByRef cityBonuses As List(Of String))
        If Not HSHPPersonal.IsEnabled() Then Return

        Dim atlas = InstanceProvider.GetInstance(Of BuildingDetailsAtlas)()

        ' Build list of all raw goods (buildings with no monthlyGoodCost prerequisites)
        Dim rawGoods As New List(Of String)()
        For Each good In zip.lexy.tgame.constants.Goods.ALL
            Dim buildingType = BuildingTypes.ForGood(good)
            If buildingType >= 1 Then
                Dim details = atlas.Get(buildingType)
                If details.monthlyGoodCost Is Nothing OrElse details.monthlyGoodCost.Count = 0 Then
                    rawGoods.Add(good)
                End If
            End If
        Next

        ' Assign 1 raw good based on city id for deterministic distribution
        If rawGoods.Count > 0 Then
            cityBonuses = New List(Of String) From {rawGoods(id Mod rawGoods.Count)}
        End If
    End Sub
End Class

' Patch SetAdditionalDifficultyModifiers
' If the "starting funds halved" modifier (id 2) is active, divide by 2 instead
' of the original behavior which sets balance to 2000.
<HarmonyPatch(GetType(GameState), "SetAdditionalDifficultyModifiers", MethodType.Normal)>
Public Class StartingFundsPatch
    Public Shared Sub Postfix(ByVal __instance As GameState)
        If Not HSHPDifficulty.IsHardMode() Then Return

        __instance.human.balance = 50000L
        If __instance.difficultyModifiers.Contains(2) Then
            __instance.human.balance = __instance.human.balance \ 2L
        End If
    End Sub
End Class

' ============================================================
' GIVE_BONUS MISSION - Increase costs 5x and set duration to 9999
' ============================================================

' Patch GoalGiver.GiveGoal to multiply give_bonus quest costs.
' Original: goods = 30/WeightMultiplier, coins = 3500
' Modified: goods x5, coins x3
<HarmonyPatch(GetType(GoalGiver), "GiveGoal", MethodType.Normal)>
Public Class GiveBonusCostPatch
    Public Shared Sub Postfix(ByRef __result As MayorGoal)
        If Not HSHPPersonal.IsEnabled() Then Return
        If __result.type <> "give_bonus" Then Return
        ApplyBonusCostMultiplier(__result)
    End Sub

    Public Shared Sub ApplyBonusCostMultiplier(ByVal goal As MayorGoal)
        For Each quest As MayorQuest In goal.quests
            If quest.type = "get_resource" Then
                ' Multiply goods amount by 5
                Dim parts = quest.args.Split("|"c)
                For j = 0 To parts.Length - 1
                    If parts(j).StartsWith("amount:") Then
                        Dim amt = Integer.Parse(parts(j).Substring(7))
                        parts(j) = "amount:" & (amt * 5).ToString()
                    End If
                Next
                quest.args = String.Join("|", parts)
            ElseIf quest.type = "pay_coins" Then
                ' Multiply coins by 3
                Dim parts = quest.args.Split("|"c)
                For j = 0 To parts.Length - 1
                    If parts(j).StartsWith("amount:") Then
                        Dim amt = Integer.Parse(parts(j).Substring(7))
                        parts(j) = "amount:" & (amt * 3).ToString()
                    End If
                Next
                quest.args = String.Join("|", parts)
            End If
        Next
    End Sub
End Class

' Patch GoalGiver.SetupBonusGoalFromGoodType to apply cost multiplier
' when selecting give_bonus from the mayor goal selection screen.
<HarmonyPatch(GetType(GoalGiver), "SetupBonusGoalFromGoodType", MethodType.Normal)>
Public Class SetupBonusCostPatch
    Public Shared Sub Postfix(ByRef __result As MayorGoal)
        If Not HSHPPersonal.IsEnabled() Then Return
        GiveBonusCostPatch.ApplyBonusCostMultiplier(__result)
    End Sub
End Class

' Patch MayorGoalImplementation.GiveTempBonus to set bonus duration to 9999 turns.
' Original duration: 204 turns (~2 years). New duration: 9999 turns (effectively permanent).
<HarmonyPatch(GetType(MayorGoalImplementation), "GiveTempBonus", MethodType.Normal)>
Public Class GiveBonusDurationPatch
    Public Shared Sub Postfix(ByVal city As City)
        If Not HSHPPersonal.IsEnabled() Then Return
        ' Find the most recently added temporary-good-bonus effect and set its duration
        For i = city.effects.Count - 1 To 0 Step -1
            If city.effects(i).type = "temporary-good-bonus" Then
                city.effects(i).remainingTurns = 9999
                Exit For
            End If
        Next
    End Sub
End Class

' Patch MayorGoalListItem.Initialize to hide build_house/build_shop entries
' in the mayor goal selection screen when Build Shops & Dwellings is enabled.
<HarmonyPatch(GetType(MayorGoalListItem), "Initialize", MethodType.Normal)>
Public Class HideBuildGoalsInSelectionPatch
    Public Shared Sub Postfix(ByVal __instance As MayorGoalListItem, ByVal goalType As String)
        If Not HSHPBuildMode.IsEnabled() Then Return
        If BlockedGoals.Blocked.Contains(goalType) Then
            __instance.gameObject.SetActive(False)
        End If
    End Sub
End Class

' Patch Localization.ForKey for HSHP info quest display text and bonus mission rename.
<HarmonyPatch(GetType(Localization), "ForKey", MethodType.Normal)>
Public Class RenameBonusMissionPatch
    Public Shared Function Prefix(ByVal key As String, ByRef __result As String) As Boolean
        ' HSHP info quest localization (not gated by any specific mode)
        Select Case key
            Case "hshp-info-title"
                __result = "HSHP Mod Features Active"
                Return False
            Case "hshp-info-text"
                __result = BuildHSHPModDescription()
                Return False
            Case "hshp-info-ok"
                __result = "Acknowledged"
                Return False
        End Select

        ' FB's Changes-only localization changes
        If Not HSHPPersonal.IsEnabled() Then Return True
        If key = "goal_give_bonus" Then
            __result = "Offer a bonus to a good"
            Return False
        End If
        Return True
    End Function

    Private Shared Function BuildHSHPModDescription() As String
        Dim lines As New List(Of String)

        If HSHPDifficulty.IsHardMode() Then
            lines.Add("<b>Black Death</b>: " & vbLf & "- All cities start with zero population with no dwellings and businesses pre-built." & vbLf & "- No AI traders." & vbLf & "- Player start with extra money, and 250 stockpile of all goods plus extra building materials in home city." & vbLf & "- add 100 repulation and 50 mayor respect per dwelling placed.")
        End If

        If HSHPBuildMode.IsEnabled() Then
            lines.Add("<b>Build Shops & Dwellings</b>: " & vbLf & "- Unlock dwelling and shop construction in the build menu." & vbLf & "- No build dwellings  & build shop mayor quests." & vbLf & "- Shops takes building slots, but protected from loss even if neglected.")
        End If

        If HSHPPersonal.IsEnabled() Then
            lines.Add("<b>FB's Changes</b>: " & vbLf & "- Businesses can only be built in towns with a bonus." & vbLf & "- Cities spawns with only one bonus good (raw only, no processed goods) but can get additional thru mayor missions." & vbLf & "- 5x the cost of bonus to goods mayor mission but made duration permanent." & vbLf & "- Population capped at dwelling capacity, details in Trader's Hall." & vbLf & "- Occupancy-based happiness instead of homelessness." & vbLf & "- Lowered reputation thresholds with 2x gain, also gain mayor respect from selling goods." & vbLf & "- No mayor elections." & vbLf & "- AI does not auto complete mayor missions.")
        End If

        If lines.Count = 0 Then Return "HSHP Unleashed Mod is active."
        Return "<b>HSHP Unleashed Mod Features Active:</b>" & vbLf & vbLf & String.Join(vbLf & vbLf, lines)
    End Function
End Class

' ============================================================
' SHOP DEMOLITION - Allow players to demolish shop buildings
' ============================================================

' Patch DetailedBuildingUi.SetupUi to show the demolish button for shops.
' The original code hides startProductionBtn.parent.parent (the button section)
' for player-managed shops. We re-activate it and hide only the irrelevant buttons,
' leaving the demolish button visible.
<HarmonyPatch(GetType(DetailedBuildingUi), "SetupUi", MethodType.Normal)>
Public Class ShopDemolishButtonPatch
    Public Shared Sub Postfix(ByVal __instance As DetailedBuildingUi)
        If Not HSHPBuildMode.IsEnabled() Then Return
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim uiType = GetType(DetailedBuildingUi)

        Dim building = uiType.GetField("building", flags).GetValue(__instance)
        If building Is Nothing Then Return
        Dim buildingType = CInt(building.GetType().GetField("type").GetValue(building))
        If Not BuildingTypes.IsShop(buildingType) Then Return

        ' Re-activate the button section container that was hidden for shops
        Dim startBtn = DirectCast(uiType.GetField("startProductionBtn", flags).GetValue(__instance), Transform)
        startBtn.parent.parent.gameObject.SetActive(True)

        ' Hide start/stop production buttons (not applicable for shops)
        startBtn.gameObject.SetActive(False)
        DirectCast(uiType.GetField("stopProductionBtn", flags).GetValue(__instance), Transform).gameObject.SetActive(False)

        ' Hide sell buttons (mod shops are not sellable)
        DirectCast(uiType.GetField("sellBusiness", flags).GetValue(__instance), Transform).gameObject.SetActive(False)
        DirectCast(uiType.GetField("businessOnSale", flags).GetValue(__instance), Transform).gameObject.SetActive(False)
    End Sub
End Class

' Patch BuildingDemolitionConfirmationWindow.ConfirmDemolish to also remove
' the Shop entry from city.shops when a shop building is demolished.
' Uses Prefix so the shop is cleaned up before the building is removed.
<HarmonyPatch(GetType(BuildingDemolitionConfirmationWindow), "ConfirmDemolish", MethodType.Normal)>
Public Class ShopDemolishCleanupPatch
    Public Shared Sub Prefix(ByVal __instance As Object)
        If Not HSHPBuildMode.IsEnabled() Then Return
        Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim building = __instance.GetType().GetField("building", flags).GetValue(__instance)
        If building Is Nothing Then Return

        Dim buildingType = CInt(building.GetType().GetField("type").GetValue(building))
        If Not BuildingTypes.IsShop(buildingType) Then Return

        Dim buildingId = CInt(building.GetType().GetField("id").GetValue(building))
        Dim gameState = InstanceProvider.GetInstance(Of GameState)()
        gameState.viewCity.shops.RemoveAll(Function(s) s.building.id = buildingId)
    End Sub
End Class

' ============================================================
' DIFFICULTY TOGGLE - Custom HSHP Hard Mode on new game screen
' ============================================================

' Module to track Black Death difficulty state and provide runtime helper
Public Module HSHPDifficulty
    Public Const MODIFIER_ID As Integer = 101

    Public Function IsHardMode() As Boolean
        Try
            Dim gs = InstanceProvider.GetInstance(Of GameState)()
            If gs IsNot Nothing AndAlso gs.difficultyModifiers IsNot Nothing Then
                If gs.difficultyModifiers.Contains(MODIFIER_ID) Then Return True
            End If
        Catch
        End Try
        Try
            Dim modsStr = "," & PlayerPrefs.GetString("difficultyModifiers", "") & ","
            If modsStr.Contains("," & MODIFIER_ID.ToString() & ",") Then Return True
        Catch
        End Try
        Return False
    End Function
End Module

' Module to track Build Shops & Dwellings feature state
Public Module HSHPBuildMode
    Public Const MODIFIER_ID As Integer = 102

    Public Function IsEnabled() As Boolean
        Try
            Dim gs = InstanceProvider.GetInstance(Of GameState)()
            If gs IsNot Nothing AndAlso gs.difficultyModifiers IsNot Nothing Then
                If gs.difficultyModifiers.Contains(MODIFIER_ID) Then Return True
            End If
        Catch
        End Try
        Try
            Dim modsStr = "," & PlayerPrefs.GetString("difficultyModifiers", "") & ","
            If modsStr.Contains("," & MODIFIER_ID.ToString() & ",") Then Return True
        Catch
        End Try
        Return False
    End Function
End Module

' Module to track FB's Changes feature state
Public Module HSHPPersonal
    Public Const MODIFIER_ID As Integer = 103

    Public Function IsEnabled() As Boolean
        Try
            Dim gs = InstanceProvider.GetInstance(Of GameState)()
            If gs IsNot Nothing AndAlso gs.difficultyModifiers IsNot Nothing Then
                If gs.difficultyModifiers.Contains(MODIFIER_ID) Then Return True
            End If
        Catch
        End Try
        Try
            Dim modsStr = "," & PlayerPrefs.GetString("difficultyModifiers", "") & ","
            If modsStr.Contains("," & MODIFIER_ID.ToString() & ",") Then Return True
        Catch
        End Try
        Return False
    End Function
End Module

' Patch NewGameMenu.SwitchToFreePlay to add HSHP mod checkboxes by cloning
' the existing "Remove Winter" checkbox. Creates 3 checkboxes:
' Build Shops & Dwellings, FB's Changes, Black Death.
' Uses a duplicate guard (finds existing clone by name) to prevent multiple clones
' if the user switches between Campaign and Free Play modes repeatedly.
<HarmonyPatch(GetType(NewGameMenu), "SwitchToFreePlay", MethodType.Normal)>
Public Class NewGameDifficultyTogglePatch
    Public Shared Sub Postfix(ByVal __instance As NewGameMenu)
        Dim pflags = BindingFlags.NonPublic Or BindingFlags.Instance

        Dim rwField = GetType(NewGameMenu).GetField("removeWinter", pflags)
        If rwField Is Nothing Then Return

        Dim sourceCheckbox = DirectCast(rwField.GetValue(__instance), Checkbox)
        If sourceCheckbox Is Nothing Then Return

        Dim parent = sourceCheckbox.transform.parent

        ' Guard: don't create duplicate if already exists
        If parent.Find("HSHP-BuildMode") IsNot Nothing Then Return

        ' Save original child positions before adding clones (layout may redistribute)
        Dim origLocalPositions As New Dictionary(Of Transform, Vector3)()
        For idx = 0 To parent.childCount - 1
            Dim child = parent.GetChild(idx)
            origLocalPositions(child) = child.localPosition
        Next

        Dim onClickField = GetType(Checkbox).GetField("onClick", pflags)

        ' --- Checkbox 1: Build Shops & Dwellings (top) ---
        Dim clone1 = UnityEngine.Object.Instantiate(sourceCheckbox.gameObject, parent)
        clone1.name = "HSHP-BuildMode"
        clone1.SetActive(True)

        Dim checkbox1 = clone1.GetComponent(Of Checkbox)()
        If checkbox1 Is Nothing Then Return

        If onClickField IsNot Nothing Then
            Dim onClickEvt1 = onClickField.GetValue(checkbox1)
            If onClickEvt1 IsNot Nothing Then
                Dim pcGroupField1 = GetType(UnityEventBase).GetField("m_PersistentCalls", pflags)
                If pcGroupField1 IsNot Nothing Then
                    Dim pcGroup1 = pcGroupField1.GetValue(DirectCast(onClickEvt1, UnityEventBase))
                    If pcGroup1 IsNot Nothing Then
                        Dim callsField1 = pcGroup1.GetType().GetField("m_Calls", pflags)
                        If callsField1 IsNot Nothing Then
                            Dim callsList1 = TryCast(callsField1.GetValue(pcGroup1), System.Collections.IList)
                            If callsList1 IsNot Nothing Then callsList1.Clear()
                        End If
                    End If
                End If
            End If
        End If

        checkbox1.Initialize("Build Shops & Dwellings", False)
        clone1.transform.SetSiblingIndex(sourceCheckbox.transform.GetSiblingIndex() + 1)

        ' --- Checkbox 2: FB's Changes (middle) ---
        Dim clone2 = UnityEngine.Object.Instantiate(sourceCheckbox.gameObject, parent)
        clone2.name = "HSHP-Personal"
        clone2.SetActive(True)

        Dim checkbox2 = clone2.GetComponent(Of Checkbox)()
        If checkbox2 IsNot Nothing Then
            If onClickField IsNot Nothing Then
                Dim onClickEvt2 = onClickField.GetValue(checkbox2)
                If onClickEvt2 IsNot Nothing Then
                    Dim pcGroupField2 = GetType(UnityEventBase).GetField("m_PersistentCalls", pflags)
                    If pcGroupField2 IsNot Nothing Then
                        Dim pcGroup2 = pcGroupField2.GetValue(DirectCast(onClickEvt2, UnityEventBase))
                        If pcGroup2 IsNot Nothing Then
                            Dim callsField2 = pcGroup2.GetType().GetField("m_Calls", pflags)
                            If callsField2 IsNot Nothing Then
                                Dim callsList2 = TryCast(callsField2.GetValue(pcGroup2), System.Collections.IList)
                                If callsList2 IsNot Nothing Then callsList2.Clear()
                            End If
                        End If
                    End If
                End If
            End If

            checkbox2.Initialize("FB's Changes", False)
            clone2.transform.SetSiblingIndex(clone1.transform.GetSiblingIndex() + 1)
        End If

        ' --- Checkbox 3: Black Death (bottom) ---
        Dim clone3 = UnityEngine.Object.Instantiate(sourceCheckbox.gameObject, parent)
        clone3.name = "HSHP-BlackDeath"
        clone3.SetActive(True)

        Dim checkbox3 = clone3.GetComponent(Of Checkbox)()
        If checkbox3 IsNot Nothing Then
            If onClickField IsNot Nothing Then
                Dim onClickEvt3 = onClickField.GetValue(checkbox3)
                If onClickEvt3 IsNot Nothing Then
                    Dim pcGroupField3 = GetType(UnityEventBase).GetField("m_PersistentCalls", pflags)
                    If pcGroupField3 IsNot Nothing Then
                        Dim pcGroup3 = pcGroupField3.GetValue(DirectCast(onClickEvt3, UnityEventBase))
                        If pcGroup3 IsNot Nothing Then
                            Dim callsField3 = pcGroup3.GetType().GetField("m_Calls", pflags)
                            If callsField3 IsNot Nothing Then
                                Dim callsList3 = TryCast(callsField3.GetValue(pcGroup3), System.Collections.IList)
                                If callsList3 IsNot Nothing Then callsList3.Clear()
                            End If
                        End If
                    End If
                End If
            End If

            checkbox3.Initialize("Black Death", False)
            clone3.transform.SetSiblingIndex(clone2.transform.GetSiblingIndex() + 1)
        End If

        ' Disable parent auto-layout so our manual positioning sticks
        Dim csf = parent.GetComponent(Of ContentSizeFitter)()
        If csf IsNot Nothing Then csf.enabled = False
        Dim vlg = parent.GetComponent(Of VerticalLayoutGroup)()
        If vlg IsNot Nothing Then vlg.enabled = False
        Dim hlg = parent.GetComponent(Of HorizontalLayoutGroup)()
        If hlg IsNot Nothing Then hlg.enabled = False

        ' Restore original positions of all existing children
        For Each kvp In origLocalPositions
            kvp.Key.localPosition = kvp.Value
        Next

        ' Position clones using localPosition
        Dim sourceRect = sourceCheckbox.GetComponent(Of RectTransform)()
        Dim srcPos = sourceCheckbox.transform.localPosition

        Dim sourceWidth = If(sourceRect IsNot Nothing, sourceRect.rect.width, 200.0F)
        If sourceWidth < 10.0F Then sourceWidth = 200.0F
        Dim xOffset = sourceWidth + 10.0F

        Dim sourceHeight = If(sourceRect IsNot Nothing, sourceRect.rect.height, 30.0F)
        If sourceHeight < 5.0F Then sourceHeight = 30.0F
        Dim rowGap = sourceHeight + 20.0F
        Dim baseY = srcPos.y + 88.0F

        ' Build Shops & Dwellings: top row
        Dim cloneRect1 = clone1.GetComponent(Of RectTransform)()
        If sourceRect IsNot Nothing AndAlso cloneRect1 IsNot Nothing Then
            cloneRect1.sizeDelta = sourceRect.sizeDelta
        End If
        clone1.transform.localPosition = New Vector3(srcPos.x + xOffset, baseY, srcPos.z)

        ' FB's Changes: second row
        If checkbox2 IsNot Nothing Then
            Dim cloneRect2 = clone2.GetComponent(Of RectTransform)()
            If sourceRect IsNot Nothing AndAlso cloneRect2 IsNot Nothing Then
                cloneRect2.sizeDelta = sourceRect.sizeDelta
            End If
            clone2.transform.localPosition = New Vector3(srcPos.x + xOffset, baseY - rowGap, srcPos.z)
        End If

        ' Black Death: third row
        If checkbox3 IsNot Nothing Then
            Dim cloneRect3 = clone3.GetComponent(Of RectTransform)()
            If sourceRect IsNot Nothing AndAlso cloneRect3 IsNot Nothing Then
                cloneRect3.sizeDelta = sourceRect.sizeDelta
            End If
            clone3.transform.localPosition = New Vector3(srcPos.x + xOffset, baseY - 2.0F * rowGap, srcPos.z)
        End If
    End Sub
End Class

' Patch NewGameMenu.SwitchToCampaign to reset our checkbox state
<HarmonyPatch(GetType(NewGameMenu), "SwitchToCampaign", MethodType.Normal)>
Public Class SwitchToCampaignResetPatch
    Public Shared Sub Postfix(ByVal __instance As NewGameMenu)
        Dim pflags = BindingFlags.NonPublic Or BindingFlags.Instance
        Dim rwField = GetType(NewGameMenu).GetField("removeWinter", pflags)
        If rwField Is Nothing Then Return

        Dim sourceCheckbox = DirectCast(rwField.GetValue(__instance), Checkbox)
        If sourceCheckbox Is Nothing Then Return

        Dim bmTransform = sourceCheckbox.transform.parent.Find("HSHP-BuildMode")
        If bmTransform IsNot Nothing Then
            Dim bmCheckbox = bmTransform.GetComponent(Of Checkbox)()
            If bmCheckbox IsNot Nothing Then bmCheckbox.Initialize("Build Shops & Dwellings", False)
        End If

        Dim pTransform = sourceCheckbox.transform.parent.Find("HSHP-Personal")
        If pTransform IsNot Nothing Then
            Dim pCheckbox = pTransform.GetComponent(Of Checkbox)()
            If pCheckbox IsNot Nothing Then pCheckbox.Initialize("FB's Changes", False)
        End If

        Dim bdTransform = sourceCheckbox.transform.parent.Find("HSHP-BlackDeath")
        If bdTransform IsNot Nothing Then
            Dim bdCheckbox = bdTransform.GetComponent(Of Checkbox)()
            If bdCheckbox IsNot Nothing Then bdCheckbox.Initialize("Black Death", False)
        End If
    End Sub
End Class

' Patch NewGameMenu.MakeNewGame to inject HSHP modifier IDs
' into difficultyModifiers before saving to PlayerPrefs.
<HarmonyPatch(GetType(NewGameMenu), "MakeNewGame", MethodType.Normal)>
Public Class MakeNewGameBlackDeathPatch
    Public Shared Sub Prefix(ByVal __instance As NewGameMenu)
        Dim pflags = BindingFlags.NonPublic Or BindingFlags.Instance

        Dim rwField = GetType(NewGameMenu).GetField("removeWinter", pflags)
        If rwField Is Nothing Then Return

        Dim sourceCheckbox = DirectCast(rwField.GetValue(__instance), Checkbox)
        If sourceCheckbox Is Nothing Then Return

        ' Access private difficultyModifiers list
        Dim modsField = GetType(NewGameMenu).GetField("difficultyModifiers", pflags)
        If modsField Is Nothing Then Return

        Dim mods = DirectCast(modsField.GetValue(__instance), List(Of Integer))
        If mods Is Nothing Then Return

        ' Read Build Shops & Dwellings checkbox
        Dim bmTransform = sourceCheckbox.transform.parent.Find("HSHP-BuildMode")
        If bmTransform IsNot Nothing Then
            Dim bmCheckbox = bmTransform.GetComponent(Of Checkbox)()
            If bmCheckbox IsNot Nothing Then
                If bmCheckbox.check Then
                    If Not mods.Contains(HSHPBuildMode.MODIFIER_ID) Then
                        mods.Add(HSHPBuildMode.MODIFIER_ID)
                    End If
                Else
                    mods.Remove(HSHPBuildMode.MODIFIER_ID)
                End If
            End If
        End If

        ' Read FB's Changes checkbox
        Dim pTransform = sourceCheckbox.transform.parent.Find("HSHP-Personal")
        If pTransform IsNot Nothing Then
            Dim pCheckbox = pTransform.GetComponent(Of Checkbox)()
            If pCheckbox IsNot Nothing Then
                If pCheckbox.check Then
                    If Not mods.Contains(HSHPPersonal.MODIFIER_ID) Then
                        mods.Add(HSHPPersonal.MODIFIER_ID)
                    End If
                Else
                    mods.Remove(HSHPPersonal.MODIFIER_ID)
                End If
            End If
        End If

        ' Read Black Death checkbox
        Dim bdTransform = sourceCheckbox.transform.parent.Find("HSHP-BlackDeath")
        If bdTransform IsNot Nothing Then
            Dim bdCheckbox = bdTransform.GetComponent(Of Checkbox)()
            If bdCheckbox IsNot Nothing Then
                If bdCheckbox.check Then
                    If Not mods.Contains(HSHPDifficulty.MODIFIER_ID) Then
                        mods.Add(HSHPDifficulty.MODIFIER_ID)
                    End If
                Else
                    mods.Remove(HSHPDifficulty.MODIFIER_ID)
                End If
            End If
        End If
    End Sub
End Class

' ============================================================
' HSHP INFO QUEST - Show mod feature descriptions via quest popup
' ============================================================

' Persistent MonoBehaviour that registers our custom QuestSO in the QuestAtlas
' as soon as the atlas is available. Must run before any quest processing.
' Uses atlas dict check instead of boolean flag so re-registration happens
' automatically when QuestAtlas is recreated on scene reload (new game).
Public Class HSHPQuestRegistrar
    Inherits MonoBehaviour

    Private Shared _instance As HSHPQuestRegistrar

    Public Shared Sub Init()
        If _instance IsNot Nothing Then Return
        Dim go = New GameObject("HSHPQuestRegistrar")
        UnityEngine.Object.DontDestroyOnLoad(go)
        _instance = go.AddComponent(Of HSHPQuestRegistrar)()
    End Sub

    ' Ensures "hshp-info" QuestSO exists in the current QuestAtlas dict.
    ' Safe to call multiple times; checks dict before adding.
    ' Returns True if registration is confirmed, False if atlas not ready.
    Public Shared Function EnsureQuestRegistered() As Boolean
        Try
            Dim questAtlas = InstanceProvider.GetInstance(Of QuestAtlas)()
            If questAtlas Is Nothing Then Return False

            Dim atlasBaseType = questAtlas.GetType().BaseType
            Dim dictField = atlasBaseType.GetField("dict", BindingFlags.NonPublic Or BindingFlags.Instance)
            If dictField Is Nothing Then Return False

            Dim dict = dictField.GetValue(questAtlas)
            If dict Is Nothing Then Return False

            Dim containsKey = dict.GetType().GetMethod("ContainsKey")
            If CBool(containsKey.Invoke(dict, New Object() {"hshp-info"})) Then Return True

            ' Get welcome QuestSO to clone its portrait (guaranteed valid)
            Dim welcomeQSO = questAtlas.Get("welcome")

            ' Create our custom QuestSO
            Dim hshpQSO = ScriptableObject.CreateInstance(Of QuestSO)()
            hshpQSO.type = "hshp-info"
            hshpQSO.portrait = welcomeQSO.portrait
            hshpQSO.title = "hshp-info-title"
            hshpQSO.text = "hshp-info-text"
            hshpQSO.option1Text = "hshp-info-ok"
            hshpQSO.option2Text = ""
            hshpQSO.option3Text = ""
            hshpQSO.option4Text = ""

            Dim addMethod = dict.GetType().GetMethod("Add")
            addMethod.Invoke(dict, New Object() {"hshp-info", hshpQSO})
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub Update()
        EnsureQuestRegistered()
    End Sub
End Class

' Patch QuestGiver.OnTurnUpdated to inject HSHP info quest at turn 1.
' Follows the same pattern as the D7 ship-loan-notification quest:
' adds quest to queuedQuests which gets included in the turn's quest list,
' then dispatched via OnRequestAddNewTurnQuests for notification + display.
<HarmonyPatch(GetType(QuestGiver), "OnTurnUpdated", MethodType.Normal)>
Public Class HSHPInfoQuestPatch
    Public Shared Sub Prefix(ByVal __instance As Object)
        Try
            ' Ensure QuestSO is registered in current atlas before quest dispatch
            HSHPQuestRegistrar.EnsureQuestRegistered()

            Dim gs = InstanceProvider.GetInstance(Of GameState)()
            If gs Is Nothing Then Return
            If gs.turn > 2 Then Return

            ' Only show when at least one HSHP mod feature is active
            If Not HSHPDifficulty.IsHardMode() AndAlso Not HSHPBuildMode.IsEnabled() AndAlso Not HSHPPersonal.IsEnabled() Then Return

            ' Don't create duplicate if already in active quests
            If gs.activeQuests IsNot Nothing AndAlso gs.activeQuests.Any(Function(q) q.type = "hshp-info") Then Return

            ' Create the quest
            Dim quest = New Quest(gs.turn)
            quest.type = "hshp-info"
            quest.args = ""

            ' Add to queuedQuests so it's included in this turn's quest list
            Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance
            Dim queuedField = __instance.GetType().GetField("queuedQuests", flags)
            If queuedField IsNot Nothing Then
                Dim queuedQuests = queuedField.GetValue(__instance)
                queuedQuests.GetType().GetMethod("Add").Invoke(queuedQuests, New Object() {quest})
            End If
        Catch
        End Try
    End Sub
End Class

' Patch Difficulty.GetLabel to return our custom label for HSHP modifier IDs
<HarmonyPatch(GetType(Difficulty), "GetLabel", MethodType.Normal)>
Public Class DifficultyLabelPatch
    Public Shared Function Prefix(ByVal difficulty As Integer, ByRef __result As String) As Boolean
        If difficulty = HSHPDifficulty.MODIFIER_ID Then
            __result = "Black Death: All cities start empty. No AI traders. 250 stockpile + dwelling materials."
            Return False
        End If
        If difficulty = HSHPBuildMode.MODIFIER_ID Then
            __result = "Build Shops & Dwellings: Unlock shop and dwelling construction in the build menu."
            Return False
        End If
        If difficulty = HSHPPersonal.MODIFIER_ID Then
            __result = "FB's Changes: Custom mayor, reputation, happiness, consumption, and mission tweaks."
            Return False
        End If
        Return True
    End Function
End Class
