# pip install opc-diag
import subprocess
import os

def ppt_to_xml(ppt_file, xml_dir):
    """Converts a PPT/PPTX file to XML files by extracting its contents."""
    if not os.path.exists(xml_dir):
        os.makedirs(xml_dir)
    subprocess.run(['opc', 'extract', ppt_file, xml_dir], check=True)

def xml_to_ppt(xml_dir, ppt_file):
    """Repackages XML files back into a PPTX file."""
    subprocess.run(['opc', 'repackage', xml_dir, ppt_file], check=True)


# Example usage:
ppt_file = 'C:/Users/kwon/Downloads/sample2.pptx'
save_dir = './xml_template'
ppt_to_xml(ppt_file, save_dir)

# save_dir = './xml_result'
new_ppt_name = 'new.pptx'
subprocess.call(f'opc repackage {save_dir} {new_ppt_name}')


# tmp = """
# <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
#   <p:cSld>
#     <p:bg>
#       <p:bgPr>
#         <a:solidFill>
#           <a:srgbClr val="FFFFFF"/>
#         </a:solidFill>
#         <a:effectLst/>
#       </p:bgPr>
#     </p:bg>
#     <p:spTree>
#       <p:nvGrpSpPr>
#         <p:cNvPr id="1" name=""/>
#         <p:cNvGrpSpPr/>
#         <p:nvPr/>
#       </p:nvGrpSpPr>
#       <p:grpSpPr>
#         <a:xfrm>
#           <a:off x="0" y="0"/>
#           <a:ext cx="0" cy="0"/>
#           <a:chOff x="0" y="0"/>
#           <a:chExt cx="0" cy="0"/>
#         </a:xfrm>
#       </p:grpSpPr>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="2" name="Picture 2"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId2"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="16078200" y="622300"/>
#             <a:ext cx="1536700" cy="9042400"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:grpSp>
#         <p:nvGrpSpPr>
#           <p:cNvPr id="3" name="Group 3"/>
#           <p:cNvGrpSpPr/>
#           <p:nvPr/>
#         </p:nvGrpSpPr>
#         <p:grpSpPr>
#           <a:xfrm>
#             <a:off x="2147483647" y="2147483647"/>
#             <a:ext cx="2147483647" cy="2147483647"/>
#             <a:chOff x="0" y="0"/>
#             <a:chExt cx="0" cy="0"/>
#           </a:xfrm>
#         </p:grpSpPr>
#       </p:grpSp>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="4" name="Picture 4"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId3"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="16200000">
#             <a:off x="16408400" y="1638300"/>
#             <a:ext cx="2222500" cy="203200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="5" name="Picture 5"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId3"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="16200000">
#             <a:off x="16408400" y="3911600"/>
#             <a:ext cx="2222500" cy="203200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="6" name="Picture 6"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId3"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="16200000">
#             <a:off x="16408400" y="6184900"/>
#             <a:ext cx="2222500" cy="203200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="7" name="Picture 7"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId3"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="16200000">
#             <a:off x="16408400" y="8445500"/>
#             <a:ext cx="2222500" cy="203200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="8" name="Picture 8"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId4"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="5400000">
#             <a:off x="16090900" y="2997200"/>
#             <a:ext cx="1143000" cy="863600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="9" name="Picture 9"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId5"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="5400000">
#             <a:off x="16090900" y="1879600"/>
#             <a:ext cx="1143000" cy="863600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="10" name="Picture 10"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId6"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm rot="5400000">
#             <a:off x="16090900" y="762000"/>
#             <a:ext cx="1143000" cy="863600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="11" name="Picture 11"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId7"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="660400" y="622300"/>
#             <a:ext cx="15989300" cy="9042400"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="12" name="TextBox 12"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1727200" y="1485900"/>
#             <a:ext cx="4038600" cy="1524000"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="78850"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="6300" b="0" i="0" u="none" strike="noStrike" spc="300">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display SemiBold"/>
#               </a:rPr>
#               <a:t>Main</a:t>
#             </a:r>
#           </a:p>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="78850"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="6300" b="0" i="0" u="none" strike="noStrike" spc="300">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display SemiBold"/>
#               </a:rPr>
#               <a:t>Experience</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="13" name="TextBox 13"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="5613400" y="1752600"/>
#             <a:ext cx="10096500" cy="1117600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="154380"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="80000"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>주요 경험 및 프로젝트를 요약하여 보여주는 페이지입니다. 각 항목에 대한 상세 설명이 이어집니다.</a:t>
#             </a:r>
#           </a:p>
#           </p:txBody>
#       </p:sp>
#       <p:pic> <p:nvPicPr>
#           <p:cNvPr id="14" name="Picture 14"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId8"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1727200" y="4356100"/>
#             <a:ext cx="6426200" cy="12700"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic> <p:nvPicPr>
#           <p:cNvPr id="15" name="Picture 15"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId9"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1727200" y="3924300"/>
#             <a:ext cx="2019300" cy="444500"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="16" name="TextBox 16"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1981200" y="3949700"/>
#             <a:ext cx="1536700" cy="355600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="ctr">
#               <a:lnSpc>
#                 <a:spcPct val="100429"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="2000" b="0" i="1" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display Bold"/>
#               </a:rPr>
#               <a:t>2023 Project</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="17" name="TextBox 17"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="3771900" y="4838700"/>
#             <a:ext cx="4356100" cy="330200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="109559"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1900" b="0" i="0" u="none" strike="noStrike" spc="-100">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Medium"/>
#               </a:rPr>
#               <a:t>AI 기반 탄소 배출량 관리 시스템</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:pic> <p:nvPicPr>
#           <p:cNvPr id="28" name="Picture 28"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId10"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="2222500" y="5003800"/>
#             <a:ext cx="774700" cy="876300"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="26" name="TextBox 26"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="3771900" y="5295900"/>
#             <a:ext cx="4191000" cy="711200"/> </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"/> <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="140000"/> </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>개요: 기업의 탄소 배출량을 추적, 예측, 분석하여 관리가 용이하도록 하는 시스템 개발</a:t>
#             </a:r>
#           </a:p>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="140000"/> </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>역할: 팀장, Backend(Django), AI(탄소 배출량 예측), 환경 관리(Docker)</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>

#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="18" name="Picture 18"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId9"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9156700" y="4356100"/>
#             <a:ext cx="6426200" cy="12700"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="19" name="Picture 19"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId9"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9156700" y="3924300"/>
#             <a:ext cx="2019300" cy="444500"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="19" name="TextBox 19"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9410700" y="3949700"/>
#             <a:ext cx="1536700" cy="355600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="ctr">
#               <a:lnSpc>
#                 <a:spcPct val="100429"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="2000" b="0" i="1" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display Bold"/>
#               </a:rPr>
#               <a:t>Point.02</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="20" name="Picture 20"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId8"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1727200" y="6985000"/>
#             <a:ext cx="6426200" cy="12700"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="21" name="Picture 21"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId9"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1727200" y="6565900"/>
#             <a:ext cx="2019300" cy="444500"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="22" name="TextBox 22"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="1981200" y="6578600"/>
#             <a:ext cx="1536700" cy="355600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="ctr">
#               <a:lnSpc>
#                 <a:spcPct val="100429"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="2000" b="0" i="1" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display Bold"/>
#               </a:rPr>
#               <a:t>Point.03</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="23" name="Picture 23"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId8"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9156700" y="6985000"/>
#             <a:ext cx="6426200" cy="12700"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="24" name="Picture 24"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId9"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9156700" y="6565900"/>
#             <a:ext cx="2019300" cy="444500"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="25" name="TextBox 25"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9410700" y="6578600"/>
#             <a:ext cx="1536700" cy="355600"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="ctr">
#               <a:lnSpc>
#                 <a:spcPct val="100429"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="2000" b="0" i="1" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:latin typeface="Playfair Display Bold"/>
#               </a:rPr>
#               <a:t>Point.04</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="32" name="TextBox 32"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="11239500" y="5295900"/>
#             <a:ext cx="4191000" cy="711200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>프로젝트 포인트에 대한 내용을 이 곳에 입력해주세요. 프로젝트 포인트에 대한 내용을 이 곳에 입력해주세요. </a:t>
#             </a:r>
#           </a:p>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>포인트에 대한 내용을 이 곳에 입력해주세요.  </a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="33" name="TextBox 33"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="11239500" y="4826000"/>
#             <a:ext cx="4356100" cy="330200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="109559"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1900" b="0" i="0" u="none" strike="noStrike" spc="-100">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Medium"/>
#               </a:rPr>
#               <a:t>두 번째 포인트를 입력하세요</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="34" name="TextBox 34"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="3771900" y="7937500"/>
#             <a:ext cx="4191000" cy="711200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>프로젝트 포인트에 대한 내용을 이 곳에 입력해주세요. </a:t>
#             </a:r>
#           </a:p>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>포인트에 대한 내용을 이 곳에 입력해주세요.  </a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="35" name="TextBox 35"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="3771900" y="7480300"/>
#             <a:ext cx="4356100" cy="330200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="109559"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1900" b="0" i="0" u="none" strike="noStrike" spc="-100">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Medium"/>
#               </a:rPr>
#               <a:t>세 번째 포인트를 입력하세요</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="36" name="TextBox 36"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="11239500" y="7937500"/>
#             <a:ext cx="4191000" cy="711200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>프로젝트 포인트에 대한 내용을 이 곳에 입력해주세요. </a:t>
#             </a:r>
#           </a:p>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="159359"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1700" b="0" i="0" u="none" strike="noStrike" spc="-200">
#                 <a:solidFill>
#                   <a:srgbClr val="222222">
#                     <a:alpha val="67843"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>포인트에 대한 내용을 이 곳에 입력해주세요.  </a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:sp> <p:nvSpPr>
#           <p:cNvPr id="37" name="TextBox 37"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="11239500" y="7480300"/>
#             <a:ext cx="4356100" cy="330200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="109559"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1900" b="0" i="0" u="none" strike="noStrike" spc="-100">
#                 <a:solidFill>
#                   <a:srgbClr val="222222"/>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Medium"/>
#               </a:rPr>
#               <a:t>네 번째 포인트를 입력하세요</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="29" name="Picture 29"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId11"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9918700" y="4889500"/>
#             <a:ext cx="673100" cy="850900"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="30" name="Picture 30"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId12"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="2235200" y="7467600"/>
#             <a:ext cx="749300" cy="1003300"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:pic>
#         <p:nvPicPr>
#           <p:cNvPr id="31" name="Picture 31"/>
#           <p:cNvPicPr>
#             <a:picLocks noChangeAspect="1"/>
#           </p:cNvPicPr>
#           <p:nvPr/>
#         </p:nvPicPr>
#         <p:blipFill>
#           <a:blip r:embed="rId13"/>
#           <a:stretch>
#             <a:fillRect/>
#           </a:stretch>
#         </p:blipFill>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="9829800" y="7658100"/>
#             <a:ext cx="838200" cy="838200"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#       </p:pic>
#       <p:sp>
#         <p:nvSpPr>
#           <p:cNvPr id="18" name="TextBox 18"/>
#           <p:cNvSpPr txBox="1"/>
#           <p:nvPr/>
#         </p:nvSpPr>
#         <p:spPr>
#           <a:xfrm>
#             <a:off x="13220700" y="965200"/>
#             <a:ext cx="2882900" cy="215900"/>
#           </a:xfrm>
#           <a:prstGeom prst="rect">
#             <a:avLst/>
#           </a:prstGeom>
#         </p:spPr>
#         <p:txBody>
#           <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>
#           <a:lstStyle/>
#           <a:p>
#             <a:pPr lvl="0" algn="l">
#               <a:lnSpc>
#                 <a:spcPct val="124499"/>
#               </a:lnSpc>
#             </a:pPr>
#             <a:r>
#               <a:rPr lang="en-US" sz="1200" b="0" i="0" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="000000">
#                     <a:alpha val="80000"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:latin typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>* </a:t>
#             </a:r>
#             <a:r>
#               <a:rPr lang="ko-KR" sz="1200" b="0" i="0" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="000000">
#                     <a:alpha val="80000"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:ea typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>페이지 내 인물사진은 샘플 이미지입니다.</a:t>
#             </a:r>
#             <a:r>
#               <a:rPr lang="en-US" sz="1200" b="0" i="0" u="none" strike="noStrike">
#                 <a:solidFill>
#                   <a:srgbClr val="000000">
#                     <a:alpha val="80000"/>
#                   </a:srgbClr>
#                 </a:solidFill>
#                 <a:latin typeface="Pretendard Light"/>
#               </a:rPr>
#               <a:t>.</a:t>
#             </a:r>
#           </a:p>
#         </p:txBody>
#       </p:sp>
#     </p:spTree>
#   </p:cSld>
#   <p:clrMapOvr>
#     <a:masterClrMapping/>
#   </p:clrMapOvr>
# </p:sld>
# """


# # from xml.etree import ElementTree as ET

# # xml_root = ET.fromstring(tmp)
# # tree = ET.ElementTree(xml_root)
# # # xml 파일로 저장
# # RESULT_XML_PROJECT_PATH = "C:/Users/kwon/Desktop/repo/github_mcp/ppt/xml/xml_result/ppt/slides/slide{slide_idx}.xml"
# # xml_path = RESULT_XML_PROJECT_PATH.format(slide_idx=5)
# # with open(xml_path, "wb") as xml_file:
# #     tree.write(xml_file, encoding="utf-8", xml_declaration=True)

# save_dir = './xml_result'

# new_ppt_name = 'new.pptx'
# subprocess.call(f'opc repackage {save_dir} {new_ppt_name}')