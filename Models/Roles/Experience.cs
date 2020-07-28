using GalaSoft.MvvmLight;
using PetaPoco;

namespace RoleDDNG.Models.Roles
{
    public class Experience : ObservableObject
    {
        private double? _fp1 = 0;

        private double? _fp10 = 0;

        private double? _fp100 = 0;

        private double? _fp101 = 0;

        private double? _fp102 = 0;

        private double? _fp103 = 0;

        private double? _fp104 = 0;

        private double? _fp105 = 0;

        private double? _fp106 = 0;

        private double? _fp107 = 0;

        private double? _fp108 = 0;

        private double? _fp109 = 0;

        private double? _fp11 = 0;

        private double? _fp110 = 0;

        private double? _fp111 = 0;

        private double? _fp112 = 0;

        private double? _fp113 = 0;

        private double? _fp114 = 0;

        private double? _fp115 = 0;

        private double? _fp116 = 0;

        private double? _fp117 = 0;

        private double? _fp118 = 0;

        private double? _fp119 = 0;

        private double? _fp12 = 0;

        private double? _fp120 = 0;

        private double? _fp121 = 0;

        private double? _fp122 = 0;

        private double? _fp123 = 0;

        private double? _fp124 = 0;

        private double? _fp125 = 0;

        private double? _fp13 = 0;

        private double? _fp14 = 0;

        private double? _fp15 = 0;

        private double? _fp16 = 0;

        private double? _fp17 = 0;

        private double? _fp18 = 0;

        private double? _fp19 = 0;

        private double? _fp2 = 0;

        private double? _fp20 = 0;

        private double? _fp21 = 0;

        private double? _fp22 = 0;

        private double? _fp23 = 0;

        private double? _fp24 = 0;

        private double? _fp25 = 0;

        private double? _fp26 = 0;

        private double? _fp27 = 0;

        private double? _fp28 = 0;

        private double? _fp29 = 0;

        private double? _fp3 = 0;

        private double? _fp30 = 0;

        private double? _fp31 = 0;

        private double? _fp32 = 0;

        private double? _fp33 = 0;

        private double? _fp34 = 0;

        private double? _fp35 = 0;

        private double? _fp36 = 0;

        private double? _fp37 = 0;

        private double? _fp38 = 0;

        private double? _fp39 = 0;

        private double? _fp4 = 0;

        private double? _fp40 = 0;

        private double? _fp41 = 0;

        private double? _fp42 = 0;

        private double? _fp43 = 0;

        private double? _fp44 = 0;

        private double? _fp45 = 0;

        private double? _fp46 = 0;

        private double? _fp47 = 0;

        private double? _fp48 = 0;

        private double? _fp49 = 0;

        private double? _fp5 = 0;

        private double? _fp50 = 0;

        private double? _fp51 = 0;

        private double? _fp52 = 0;

        private double? _fp53 = 0;

        private double? _fp54 = 0;

        private double? _fp55 = 0;

        private double? _fp56 = 0;

        private double? _fp57 = 0;

        private double? _fp58 = 0;

        private double? _fp59 = 0;

        private double? _fp6 = 0;

        private double? _fp60 = 0;

        private double? _fp61 = 0;

        private double? _fp62 = 0;

        private double? _fp63 = 0;

        private double? _fp64 = 0;

        private double? _fp65 = 0;

        private double? _fp66 = 0;

        private double? _fp67 = 0;

        private double? _fp68 = 0;

        private double? _fp69 = 0;

        private double? _fp7 = 0;

        private double? _fp70 = 0;

        private double? _fp71 = 0;

        private double? _fp72 = 0;

        private double? _fp73 = 0;

        private double? _fp74 = 0;

        private double? _fp75 = 0;

        private double? _fp76 = 0;

        private double? _fp77 = 0;

        private double? _fp78 = 0;

        private double? _fp79 = 0;

        private double? _fp8 = 0;

        private double? _fp80 = 0;

        private double? _fp81 = 0;

        private double? _fp82 = 0;

        private double? _fp83 = 0;

        private double? _fp84 = 0;

        private double? _fp85 = 0;

        private double? _fp86 = 0;

        private double? _fp87 = 0;

        private double? _fp88 = 0;

        private double? _fp89 = 0;

        private double? _fp9 = 0;

        private double? _fp90 = 0;

        private double? _fp91 = 0;

        private double? _fp92 = 0;

        private double? _fp93 = 0;

        private double? _fp94 = 0;

        private double? _fp95 = 0;

        private double? _fp96 = 0;

        private double? _fp97 = 0;

        private double? _fp98 = 0;

        private double? _fp99 = 0;

        private double? _niveau = 0;

        public Experience()
        {
        }

        [Column("fp1")]
        public double? FP1 { get => _fp1; set { Set(nameof(FP1), ref _fp1, value); } }

        [Column("fp10")]
        public double? FP10 { get => _fp10; set { Set(nameof(FP10), ref _fp10, value); } }

        [Column("fp100")]
        public double? FP100 { get => _fp100; set { Set(nameof(FP100), ref _fp100, value); } }

        [Column("fp101")]
        public double? FP101 { get => _fp101; set { Set(nameof(FP101), ref _fp101, value); } }

        [Column("fp102")]
        public double? FP102 { get => _fp102; set { Set(nameof(FP102), ref _fp102, value); } }

        [Column("fp103")]
        public double? FP103 { get => _fp103; set { Set(nameof(FP103), ref _fp103, value); } }

        [Column("fp104")]
        public double? FP104 { get => _fp104; set { Set(nameof(FP104), ref _fp104, value); } }

        [Column("fp105")]
        public double? FP105 { get => _fp105; set { Set(nameof(FP105), ref _fp105, value); } }

        [Column("fp106")]
        public double? FP106 { get => _fp106; set { Set(nameof(FP106), ref _fp106, value); } }

        [Column("fp107")]
        public double? FP107 { get => _fp107; set { Set(nameof(FP107), ref _fp107, value); } }

        [Column("fp108")]
        public double? FP108 { get => _fp108; set { Set(nameof(FP108), ref _fp108, value); } }

        [Column("fp109")]
        public double? FP109 { get => _fp109; set { Set(nameof(FP109), ref _fp109, value); } }

        [Column("fp11")]
        public double? FP11 { get => _fp11; set { Set(nameof(FP11), ref _fp11, value); } }

        [Column("fp110")]
        public double? FP110 { get => _fp110; set { Set(nameof(FP110), ref _fp110, value); } }

        [Column("fp111")]
        public double? FP111 { get => _fp111; set { Set(nameof(FP111), ref _fp111, value); } }

        [Column("fp112")]
        public double? FP112 { get => _fp112; set { Set(nameof(FP112), ref _fp112, value); } }

        [Column("fp113")]
        public double? FP113 { get => _fp113; set { Set(nameof(FP113), ref _fp113, value); } }

        [Column("fp114")]
        public double? FP114 { get => _fp114; set { Set(nameof(FP114), ref _fp114, value); } }

        [Column("fp115")]
        public double? FP115 { get => _fp115; set { Set(nameof(FP115), ref _fp115, value); } }

        [Column("fp116")]
        public double? FP116 { get => _fp116; set { Set(nameof(FP116), ref _fp116, value); } }

        [Column("fp117")]
        public double? FP117 { get => _fp117; set { Set(nameof(FP117), ref _fp117, value); } }

        [Column("fp118")]
        public double? FP118 { get => _fp118; set { Set(nameof(FP118), ref _fp118, value); } }

        [Column("fp119")]
        public double? FP119 { get => _fp119; set { Set(nameof(FP119), ref _fp119, value); } }

        [Column("fp12")]
        public double? FP12 { get => _fp12; set { Set(nameof(FP12), ref _fp12, value); } }

        [Column("fp120")]
        public double? FP120 { get => _fp120; set { Set(nameof(FP120), ref _fp120, value); } }

        [Column("fp121")]
        public double? FP121 { get => _fp121; set { Set(nameof(FP121), ref _fp121, value); } }

        [Column("fp122")]
        public double? FP122 { get => _fp122; set { Set(nameof(FP122), ref _fp122, value); } }

        [Column("fp123")]
        public double? FP123 { get => _fp123; set { Set(nameof(FP123), ref _fp123, value); } }

        [Column("fp124")]
        public double? FP124 { get => _fp124; set { Set(nameof(FP124), ref _fp124, value); } }

        [Column("fp125")]
        public double? FP125 { get => _fp125; set { Set(nameof(FP125), ref _fp125, value); } }

        [Column("fp13")]
        public double? FP13 { get => _fp13; set { Set(nameof(FP13), ref _fp13, value); } }

        [Column("fp14")]
        public double? FP14 { get => _fp14; set { Set(nameof(FP14), ref _fp14, value); } }

        [Column("fp15")]
        public double? FP15 { get => _fp15; set { Set(nameof(FP15), ref _fp15, value); } }

        [Column("fp16")]
        public double? FP16 { get => _fp16; set { Set(nameof(FP16), ref _fp16, value); } }

        [Column("fp17")]
        public double? FP17 { get => _fp17; set { Set(nameof(FP17), ref _fp17, value); } }

        [Column("fp18")]
        public double? FP18 { get => _fp18; set { Set(nameof(FP18), ref _fp18, value); } }

        [Column("fp19")]
        public double? FP19 { get => _fp19; set { Set(nameof(FP19), ref _fp19, value); } }

        [Column("fp2")]
        public double? FP2 { get => _fp2; set { Set(nameof(FP2), ref _fp2, value); } }

        [Column("fp20")]
        public double? FP20 { get => _fp20; set { Set(nameof(FP20), ref _fp20, value); } }

        [Column("fp21")]
        public double? FP21 { get => _fp21; set { Set(nameof(FP21), ref _fp21, value); } }

        [Column("fp22")]
        public double? FP22 { get => _fp22; set { Set(nameof(FP22), ref _fp22, value); } }

        [Column("fp23")]
        public double? FP23 { get => _fp23; set { Set(nameof(FP23), ref _fp23, value); } }

        [Column("fp24")]
        public double? FP24 { get => _fp24; set { Set(nameof(FP24), ref _fp24, value); } }

        [Column("fp25")]
        public double? FP25 { get => _fp25; set { Set(nameof(FP25), ref _fp25, value); } }

        [Column("fp26")]
        public double? FP26 { get => _fp26; set { Set(nameof(FP26), ref _fp26, value); } }

        [Column("fp27")]
        public double? FP27 { get => _fp27; set { Set(nameof(FP27), ref _fp27, value); } }

        [Column("fp28")]
        public double? FP28 { get => _fp28; set { Set(nameof(FP28), ref _fp28, value); } }

        [Column("fp29")]
        public double? FP29 { get => _fp29; set { Set(nameof(FP29), ref _fp29, value); } }

        [Column("fp3")]
        public double? FP3 { get => _fp3; set { Set(nameof(FP3), ref _fp3, value); } }

        [Column("fp30")]
        public double? FP30 { get => _fp30; set { Set(nameof(FP30), ref _fp30, value); } }

        [Column("fp31")]
        public double? FP31 { get => _fp31; set { Set(nameof(FP31), ref _fp31, value); } }

        [Column("fp32")]
        public double? FP32 { get => _fp32; set { Set(nameof(FP32), ref _fp32, value); } }

        [Column("fp33")]
        public double? FP33 { get => _fp33; set { Set(nameof(FP33), ref _fp33, value); } }

        [Column("fp34")]
        public double? FP34 { get => _fp34; set { Set(nameof(FP34), ref _fp34, value); } }

        [Column("fp35")]
        public double? FP35 { get => _fp35; set { Set(nameof(FP35), ref _fp35, value); } }

        [Column("fp36")]
        public double? FP36 { get => _fp36; set { Set(nameof(FP36), ref _fp36, value); } }

        [Column("fp37")]
        public double? FP37 { get => _fp37; set { Set(nameof(FP37), ref _fp37, value); } }

        [Column("fp38")]
        public double? FP38 { get => _fp38; set { Set(nameof(FP38), ref _fp38, value); } }

        [Column("fp39")]
        public double? FP39 { get => _fp39; set { Set(nameof(FP39), ref _fp39, value); } }

        [Column("fp4")]
        public double? FP4 { get => _fp4; set { Set(nameof(FP4), ref _fp4, value); } }

        [Column("fp40")]
        public double? FP40 { get => _fp40; set { Set(nameof(FP40), ref _fp40, value); } }

        [Column("fp41")]
        public double? FP41 { get => _fp41; set { Set(nameof(FP41), ref _fp41, value); } }

        [Column("fp42")]
        public double? FP42 { get => _fp42; set { Set(nameof(FP42), ref _fp42, value); } }

        [Column("fp43")]
        public double? FP43 { get => _fp43; set { Set(nameof(FP43), ref _fp43, value); } }

        [Column("fp44")]
        public double? FP44 { get => _fp44; set { Set(nameof(FP44), ref _fp44, value); } }

        [Column("fp45")]
        public double? FP45 { get => _fp45; set { Set(nameof(FP45), ref _fp45, value); } }

        [Column("fp46")]
        public double? FP46 { get => _fp46; set { Set(nameof(FP46), ref _fp46, value); } }

        [Column("fp47")]
        public double? FP47 { get => _fp47; set { Set(nameof(FP47), ref _fp47, value); } }

        [Column("fp48")]
        public double? FP48 { get => _fp48; set { Set(nameof(FP48), ref _fp48, value); } }

        [Column("fp49")]
        public double? FP49 { get => _fp49; set { Set(nameof(FP49), ref _fp49, value); } }

        [Column("fp5")]
        public double? FP5 { get => _fp5; set { Set(nameof(FP5), ref _fp5, value); } }

        [Column("fp50")]
        public double? FP50 { get => _fp50; set { Set(nameof(FP50), ref _fp50, value); } }

        [Column("fp51")]
        public double? FP51 { get => _fp51; set { Set(nameof(FP51), ref _fp51, value); } }

        [Column("fp52")]
        public double? FP52 { get => _fp52; set { Set(nameof(FP52), ref _fp52, value); } }

        [Column("fp53")]
        public double? FP53 { get => _fp53; set { Set(nameof(FP53), ref _fp53, value); } }

        [Column("fp54")]
        public double? FP54 { get => _fp54; set { Set(nameof(FP54), ref _fp54, value); } }

        [Column("fp55")]
        public double? FP55 { get => _fp55; set { Set(nameof(FP55), ref _fp55, value); } }

        [Column("fp56")]
        public double? FP56 { get => _fp56; set { Set(nameof(FP56), ref _fp56, value); } }

        [Column("fp57")]
        public double? FP57 { get => _fp57; set { Set(nameof(FP57), ref _fp57, value); } }

        [Column("fp58")]
        public double? FP58 { get => _fp58; set { Set(nameof(FP58), ref _fp58, value); } }

        [Column("fp59")]
        public double? FP59 { get => _fp59; set { Set(nameof(FP59), ref _fp59, value); } }

        [Column("fp6")]
        public double? FP6 { get => _fp6; set { Set(nameof(FP6), ref _fp6, value); } }

        [Column("fp60")]
        public double? FP60 { get => _fp60; set { Set(nameof(FP60), ref _fp60, value); } }

        [Column("fp61")]
        public double? FP61 { get => _fp61; set { Set(nameof(FP61), ref _fp61, value); } }

        [Column("fp62")]
        public double? FP62 { get => _fp62; set { Set(nameof(FP62), ref _fp62, value); } }

        [Column("fp63")]
        public double? FP63 { get => _fp63; set { Set(nameof(FP63), ref _fp63, value); } }

        [Column("fp64")]
        public double? FP64 { get => _fp64; set { Set(nameof(FP64), ref _fp64, value); } }

        [Column("fp65")]
        public double? FP65 { get => _fp65; set { Set(nameof(FP65), ref _fp65, value); } }

        [Column("fp66")]
        public double? FP66 { get => _fp66; set { Set(nameof(FP66), ref _fp66, value); } }

        [Column("fp67")]
        public double? FP67 { get => _fp67; set { Set(nameof(FP67), ref _fp67, value); } }

        [Column("fp68")]
        public double? FP68 { get => _fp68; set { Set(nameof(FP68), ref _fp68, value); } }

        [Column("fp69")]
        public double? FP69 { get => _fp69; set { Set(nameof(FP69), ref _fp69, value); } }

        [Column("fp7")]
        public double? FP7 { get => _fp7; set { Set(nameof(FP7), ref _fp7, value); } }

        [Column("fp70")]
        public double? FP70 { get => _fp70; set { Set(nameof(FP70), ref _fp70, value); } }

        [Column("fp71")]
        public double? FP71 { get => _fp71; set { Set(nameof(FP71), ref _fp71, value); } }

        [Column("fp72")]
        public double? FP72 { get => _fp72; set { Set(nameof(FP72), ref _fp72, value); } }

        [Column("fp73")]
        public double? FP73 { get => _fp73; set { Set(nameof(FP73), ref _fp73, value); } }

        [Column("fp74")]
        public double? FP74 { get => _fp74; set { Set(nameof(FP74), ref _fp74, value); } }

        [Column("fp75")]
        public double? FP75 { get => _fp75; set { Set(nameof(FP75), ref _fp75, value); } }

        [Column("fp76")]
        public double? FP76 { get => _fp76; set { Set(nameof(FP76), ref _fp76, value); } }

        [Column("fp77")]
        public double? FP77 { get => _fp77; set { Set(nameof(FP77), ref _fp77, value); } }

        [Column("fp78")]
        public double? FP78 { get => _fp78; set { Set(nameof(FP78), ref _fp78, value); } }

        [Column("fp79")]
        public double? FP79 { get => _fp79; set { Set(nameof(FP79), ref _fp79, value); } }

        [Column("fp8")]
        public double? FP8 { get => _fp8; set { Set(nameof(FP8), ref _fp8, value); } }

        [Column("fp80")]
        public double? FP80 { get => _fp80; set { Set(nameof(FP80), ref _fp80, value); } }

        [Column("fp81")]
        public double? FP81 { get => _fp81; set { Set(nameof(FP81), ref _fp81, value); } }

        [Column("fp82")]
        public double? FP82 { get => _fp82; set { Set(nameof(FP82), ref _fp82, value); } }

        [Column("fp83")]
        public double? FP83 { get => _fp83; set { Set(nameof(FP83), ref _fp83, value); } }

        [Column("fp84")]
        public double? FP84 { get => _fp84; set { Set(nameof(FP84), ref _fp84, value); } }

        [Column("fp85")]
        public double? FP85 { get => _fp85; set { Set(nameof(FP85), ref _fp85, value); } }

        [Column("fp86")]
        public double? FP86 { get => _fp86; set { Set(nameof(FP86), ref _fp86, value); } }

        [Column("fp87")]
        public double? FP87 { get => _fp87; set { Set(nameof(FP87), ref _fp87, value); } }

        [Column("fp88")]
        public double? FP88 { get => _fp88; set { Set(nameof(FP88), ref _fp88, value); } }

        [Column("fp89")]
        public double? FP89 { get => _fp89; set { Set(nameof(FP89), ref _fp89, value); } }

        [Column("fp9")]
        public double? FP9 { get => _fp9; set { Set(nameof(FP9), ref _fp9, value); } }

        [Column("fp90")]
        public double? FP90 { get => _fp90; set { Set(nameof(FP90), ref _fp90, value); } }

        [Column("fp91")]
        public double? FP91 { get => _fp91; set { Set(nameof(FP91), ref _fp91, value); } }

        [Column("fp92")]
        public double? FP92 { get => _fp92; set { Set(nameof(FP92), ref _fp92, value); } }

        [Column("fp93")]
        public double? FP93 { get => _fp93; set { Set(nameof(FP93), ref _fp93, value); } }

        [Column("fp94")]
        public double? FP94 { get => _fp94; set { Set(nameof(FP94), ref _fp94, value); } }

        [Column("fp95")]
        public double? FP95 { get => _fp95; set { Set(nameof(FP95), ref _fp95, value); } }

        [Column("fp96")]
        public double? FP96 { get => _fp96; set { Set(nameof(FP96), ref _fp96, value); } }

        [Column("fp97")]
        public double? FP97 { get => _fp97; set { Set(nameof(FP97), ref _fp97, value); } }

        [Column("fp98")]
        public double? FP98 { get => _fp98; set { Set(nameof(FP98), ref _fp98, value); } }

        [Column("fp99")]
        public double? FP99 { get => _fp99; set { Set(nameof(FP99), ref _fp99, value); } }

        [Column("niveau")]
        public double? Niveau { get => _niveau; set { Set(nameof(Niveau), ref _niveau, value); } }
    }
}