﻿'------------------------------------------------------------------------------
' <auto-generated>
'     這段程式碼是由工具產生的。
'     執行階段版本:4.0.30319.18444
'
'     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
'     變更將會遺失。
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    '這個類別是自動產生的，是利用 StronglyTypedResourceBuilder
    '類別透過 ResGen 或 Visual Studio 這類工具。
    '若要加入或移除成員，請編輯您的 .ResX 檔，然後重新執行 ResGen
    '(利用 /str 選項)，或重建您的 VS 專案。
    '''<summary>
    '''  用於查詢當地語系化字串等的強類型資源類別。
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  傳回這個類別使用的快取的 ResourceManager 執行個體。
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Chao.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  覆寫目前執行緒的 CurrentUICulture 屬性，對象是所有
        '''  使用這個強類型資源類別的資源查閱。
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  查詢類似 最大轉速穏定運轉測試時散熱風扇必須有70 %以上時間維持運轉(油門到底) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A1_step1() As String
            Get
                Return ResourceManager.GetString("A1_step1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 最大挖掘半徑75%距地面0.5公尺高處，切刃背面與地面呈60度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step1() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 高速怠轉 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step2() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 抓斗移動至操作範圍50%，並保持距地面0.5公尺之高度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step3() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step3", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 抓斗最大伸展高度之30% 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step4() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step4", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 吊桿向左方向旋轉90度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step5() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step5", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 最大伸展高度之60%停 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step6() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step6", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 伸展至75%並使抓斗之切刃呈垂直時傾卸(外推黃錐卸土) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step7() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step7", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 向右迴轉90度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step8() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step8", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 回致原先之位置(下降回紅錐) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Excavator_step9() As String
            Get
                Return ResourceManager.GetString("A2_Excavator_step9", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 吊桿向左方向旋轉45度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Loader_Excavator_step5() As String
            Get
                Return ResourceManager.GetString("A2_Loader_Excavator_step5", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 向右迴轉45度 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Loader_Excavator_step8() As String
            Get
                Return ResourceManager.GetString("A2_Loader_Excavator_step8", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 高速怠轉(油門到底) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Loader_step1() As String
            Get
                Return ResourceManager.GetString("A2_Loader_step1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 抓斗上舉制最高舉程之75%處(舉高X米) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Loader_step2() As String
            Get
                Return ResourceManager.GetString("A2_Loader_step2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 再回復至原始位置(放下) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A2_Loader_step3() As String
            Get
                Return ResourceManager.GetString("A2_Loader_step3", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 抓斗底距地面0.3公尺±0.05公尺(起點紅錐斗高1尺) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Loader_step1() As String
            Get
                Return ResourceManager.GetString("A3_Loader_step1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 引擎高速怠轉(高速怠轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Loader_step2() As String
            Get
                Return ResourceManager.GetString("A3_Loader_step2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 履帶式前進速度儘量接近且不超過4公里/小時，膠輪式前進速度儘量接近且不超過8公里/小時(低檔前進白錐) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Loader_step3() As String
            Get
                Return ResourceManager.GetString("A3_Loader_step3", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 後退行走與速度無關，可使用適宜之變速檔。(後退回紅錐) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Loader_step4() As String
            Get
                Return ResourceManager.GetString("A3_Loader_step4", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 排土板為標準裝置，距地面高0.3公尺±0.05公尺(起點紅錐板高1尺) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Tractor_step1() As String
            Get
                Return ResourceManager.GetString("A3_Tractor_step1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 高速怠轉(油門到底) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Tractor_step2() As String
            Get
                Return ResourceManager.GetString("A3_Tractor_step2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 膠輪式於堅硬反射面行走，前進速度儘量接近且不超過每小時8公里，履帶式及鋼輪式於砂土上行走，前進速度儘量接近且不超過每小時4公里(前進至白錐)(停) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Tractor_step3() As String
            Get
                Return ResourceManager.GetString("A3_Tractor_step3", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 後退速度則視情況使用變速檔。(後退回紅錐) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A3_Tractor_step4() As String
            Get
                Return ResourceManager.GetString("A3_Tractor_step4", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 定置高速怠轉狀態。(全速空轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Asphalt_Finisher() As String
            Get
                Return ResourceManager.GetString("A4_Asphalt_Finisher", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 定置高速怠轉。(全速空轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Auger_Drill_Driver() As String
            Get
                Return ResourceManager.GetString("A4_Auger_Drill_Driver", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 原則上定數回轉與定額負載之狀態。(定速定載) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Compressor() As String
            Get
                Return ResourceManager.GetString("A4_Compressor", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 以規定之工作壓力為測定作業狀態，將鑿桿(chisel)強烈押在控制板，須避免組裝部分影響測值。作業者不可站在噪音測線上 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Concrete_Breaker() As String
            Get
                Return ResourceManager.GetString("A4_Concrete_Breaker", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 以定速回轉切割混凝土，深度為刀片直徑之1/4 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Concrete_Cutter() As String
            Get
                Return ResourceManager.GetString("A4_Concrete_Cutter", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 以最大之運轉狀態壓送混凝土，此時吊桿應向水平方向延伸，配管長度約為10公尺。(最大之運轉狀態壓送混凝土) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Concrete_Pump() As String
            Get
                Return ResourceManager.GetString("A4_Concrete_Pump", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 吊桿之角度為60度，勾子抓斗等以上卷狀態，定置高速怠轉。(全速空轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Crane() As String
            Get
                Return ResourceManager.GetString("A4_Crane", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 無負載定速回轉(60Hz)。(無負載定速回轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Generator() As String
            Get
                Return ResourceManager.GetString("A4_Generator", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 可裝載道碴之機具，以裝載最大量之狀態，定置高速怠轉。(滿載全速空轉) 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Roller() As String
            Get
                Return ResourceManager.GetString("A4_Roller", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類似 吊在空中之狀態，不超過地面上50公分為原則並使其產生最大振動數 的當地語系化字串。
        '''</summary>
        Friend ReadOnly Property A4_Vibrating_Hammer() As String
            Get
                Return ResourceManager.GetString("A4_Vibrating_Hammer", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property vibrating_hammer() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("vibrating_hammer", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 全套管鑽掘機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("全套管鑽掘機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 卡車起重機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("卡車起重機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 土壤取樣器_地鑽___Earth_auger_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("土壤取樣器_地鑽___Earth_auger_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 壓路機_rollers_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("壓路機_rollers_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 小型膠輪式裝料機_compact_loader__wheeled_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("小型膠輪式裝料機_compact_loader__wheeled_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 小型開挖機_compact_excavator_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("小型開挖機_compact_excavator_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 履帶式推土機_crawler_dozer_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("履帶式推土機_crawler_dozer_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 履帶式裝料機_crawler_loader_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("履帶式裝料機_crawler_loader_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 履帶式開挖機_crawler_excavator_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("履帶式開挖機_crawler_excavator_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 履帶式開挖裝料機_crawler_backhoe_loader_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("履帶式開挖裝料機_crawler_backhoe_loader_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 履帶起重機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("履帶起重機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 拔樁機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("拔樁機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 振動打樁機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("振動打樁機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 油壓式打樁機_Hydraulic_pile_driver_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("油壓式打樁機_Hydraulic_pile_driver_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 混凝土割切機_Concrete_cutter_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("混凝土割切機_Concrete_cutter_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 混凝土泵車_Concrete_pump_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("混凝土泵車_Concrete_pump_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 混凝土破碎機_Concrete_breaker_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("混凝土破碎機_Concrete_breaker_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 滑移裝料機_skid_steer_loader_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("滑移裝料機_skid_steer_loader_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 瀝青混凝土舖築機_Asphalt_finisher_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("瀝青混凝土舖築機_Asphalt_finisher_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 發電機_Generator_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("發電機_Generator_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 空氣壓縮機_Compressor_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("空氣壓縮機_Compressor_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 膠輪式推土機_wheeled_dozer_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("膠輪式推土機_wheeled_dozer_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 膠輪式裝料機_wheeled_loader_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("膠輪式裝料機_wheeled_loader_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 膠輪式開挖機_wheeled_excavator_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("膠輪式開挖機_wheeled_excavator_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 膠輪式開挖裝料機_wheeled_backhoe_loader_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("膠輪式開挖裝料機_wheeled_backhoe_loader_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 輪式起重機() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("輪式起重機", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 鑽土機_Earth_drill_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("鑽土機_Earth_drill_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
        
        '''<summary>
        '''  查詢類型 System.Drawing.Bitmap 的當地語系化資源。
        '''</summary>
        Friend ReadOnly Property 鑽岩機_Rock_breaker_() As System.Drawing.Bitmap
            Get
                Dim obj As Object = ResourceManager.GetObject("鑽岩機_Rock_breaker_", resourceCulture)
                Return CType(obj,System.Drawing.Bitmap)
            End Get
        End Property
    End Module
End Namespace
