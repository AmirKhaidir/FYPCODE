   M S B A M L   `   `   `  ����M  IFileAutomation 1.0, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null9!clr-namespace:FileAutomation_1._0FileAutomation_1._0  5      Y UGdPicture.NET.14, Version=14.1.0.90, Culture=neutral, PublicKeyToken=f52a2e60ad468dbbK7clr-namespace:GdPicture14.WPF;assembly=GdPicture.NET.14GdPicture14.WPF 5      $    FileAutomation_1._0.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      2self!clr-namespace:FileAutomation_1._0     5      Ggdp7clr-namespace:GdPicture14.WPF;assembly=GdPicture.NET.14     5      -   5          � Title$  File Automation��5      $
��1280��6   $	��768��6*    � WindowState$ 	Maximized=�67    � WindowStartupLocation$ CenterScreen=�6O    J� CommandBindings 5      )  #System.Windows.Input.CommandBinding  5   
   -   5           Command$% self:CustomCommand.PreviousPage��5      6o     5   
   -   5        $! self:CustomCommand.NextPage��5      6g     5   
   -   5        $ self:CustomCommand.Search��5      6c     5   
   -   5        $# self:CustomCommand.SearchNext��5      6g     5   
   -   5        $' !self:CustomCommand.SearchPrevious��5      6k   .��5      �  � ColumnDefinitions 5   
   �� 5      $	��650 �6   �� 5      $��* �6   .+�5   
   � $��0��6   .+�5      �� $
��Left=�6   $	��Top=�65   .@�5      �� $
��Home��6   .��5      � ���	#FFE5E5E5 � RowDefinitions 5       �� 5!      $��70 �6,   �� 5"      $��* �6,   .+�5$      � $��0��6    5%      �� 5&   "   $��90 �63   �� 5'   "   $	��320 �63   �� 5(   "   $��* �63   .+�5*      � $��0��6#   .+�5+   "   �� -   5         ��Name$ AddFile��5+   )   $	��Top=�6;   $
��10,5q�6S   $��45��6w   ���#D1D0CE$	��0.5q�5,   %   .��5-   &   �� .+�5.   *   �� $��Images/folder.png��60   $��25��6K   $��25��6V   �� 5/   *   $��0q�64    �� TextAlignment$ Center=�6?   
New File6`   50   '   51   #   52      � 53      $��1��6#   .+�54   "   �� -   5        	 ��Name$	 searchText_tb��54   *   $
��Left=�6A   $��23��6\   $��	10,22,0,0q�6h   
 �� TextWrapping$

 Wrap=�6{   $	��Top=�6�   $	��306��6�   �� 55   "   $��Enter text to search:��6(   $
��Left=�6H   $��10,0,0,0q�6c   $	��Top=�6u   56      � 57      $��2��6#   .+�58   "   �� -	   5        $ 
searchWord��58   )   $��Search��6=   $
��Left=�6N   $��	10,21,0,0q�6i   $	��Top=�6}   $��75��6�   $��23��6�   6�   59      5;      � 5<      $��1��6   $��10,20q�6,   .+�5=      � .+�5>   "   �� -
   5         ��Name$ fileLIstView��5>   +   $	��600��6�    � ContextMenu 5?   &   �� 5@   *   .@�5A   .   m� -   5        $��Show Extracted Text��5A   7   6p   m� 5B   .   -   5         m�Name$ btn_deleteFile��5B   7   $��Delete��6O   6_   m� 5C   .   -   5        $ btn_refresh��5C   7   $��Refresh��6L   6]   5D   +    �� View 5F   &   �� 5G   *   .P�5H   .   �� $
��Name��6=    �� Width$	 150��6K    �� DisplayMemberBinding 6W   �� *Name+�� 5I   .   $
��Path��6=   $	 350��6K    6W   �� *
FilePath+�� 5J   .   $
��Type��6=   $	 100��6K    6W   �� *Type+5K   +   5M   #   5N      5P      5R      5S      �� 5T      $��
Automation��6   .��5U      �  5V      �� 5W      $	��120 �6,   �� 5X      $
��600* �6,   �� 5Y      $��30 �6,   �� 5Z      $��80 �6,   .+�5\      � $��0��6   $��10,10q�6,    5]      �� 5^   "   $	��100 �63   �� 5_   "   $	��100 �63   �� 5`   "   $	��100 �63   �� 5a   "   $	��100 �63   �� 5b   "   $	��100 �63   �� 5c   "   $	��100 �63   �� 5d   "   $	��100 �63   �� 5e   "   $	��100 �63   �� 5f   "   $	��100 �63   �� 5g   "   $��* �63   .+�5j      �� -   5        $ 	btnRecord��5k      $��0��5l      $��0,0,5,0q�5m      
������ $��0q�5o      .��5q   "   �� .+�5r   &   �� $��Images/recording.png��6,    �� BitmapScalingMode$ NearestNeighbor=�6J   $
��Fill=�5s   %   $��64��5t   %   $��64��5u   %   �� 5v   &   $��0q�60   $ Center=�6;   Record6Z   5w   #   5x      �� 5y      -   5        $ btnPlayback��5z      $��1��5{      $��0,0,5,0q�5|      
������ $��0q�5~      .��5�   "   �� .+�5�   &   �� $��Images/play.png��6,   $ NearestNeighbor=�6E   $
��Fill=�5�   %   $��64��5�   %   $��64��5�   %   �� 5�   &   $ Center=�60   Play6M   5�   #   5�      �� 5�      -   5        $ btnStop��5�      $��2��5�      $��0,0,5,0q�5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/stop-button.png��6,   $ NearestNeighbor=�6L   $
��Fill=�5�   %   $��64��5�   %   $��64��5�   %   �� 5�   &   $ Center=�60   Stop6M   5�   #   5�      �� 5�      -   5        $ btn_loadAutomation��5�      $��3��5�      $��0,0,0,0q�5�      
������ $��0q�5�      $	��106��5�   >   .��5�   "   �� .+�5�   &   �� $��Images/sync.png��6,   $ NearestNeighbor=�6E   $
��Fill=�5�   %   $��64��5�   %   $��64��5�   %   �� 5�   &   $��0q�60   $ Center=�6;   �� 6S   $�	Load save��6W   6k   5�   #   5�      �� 5�      -   5        $ btn_saveAutomation��5�      $��4��5�      $��0,0,5,0q�5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/diskette.png��6,   $ NearestNeighbor=�5�   (   $
��Fill=�5�   (   $��60��5�   (   $��60��5�   (   �� 5�   &   $ Center=�60   Save6M   5�   #   5�      �� 5�      -   5        $ btnCreateClick��5�      $��5��5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/mouse.png��6,   $ NearestNeighbor=�6F   $
��Fill=�5�   %   $��74��5�   %   $��74��5�   %   �� 5�   &   $ Center=�60   Click6N   5�   #   5�      �� 5�      -   5        $ btnInsertKey��5�      $��6��5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/keyboard.png��6,   $ NearestNeighbor=�6I   $
��Fill=�5�   %   $��74��5�   %   $��74��5�   %   �� 5�   &   $	��Key��60   $ Center=�6;   5�   #   5�      �� 5�      -   5        $ btnCreateText��5�      $��7��5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/strikethrough.png��6,   $ NearestNeighbor=�6N   $
��Fill=�5�   %   $��74��5�   %   $��74��5�   %   �� 5�   &   $
��Text��60   $ Center=�6<   5�   #   5�      �� 5�      -   5        $ btnWait��5�      $��8��5�      
������ $��0q�5�      .��5�   "   �� .+�5�   &   �� $��Images/chronometer.png��6,   $ NearestNeighbor=�6L   $
��Fill=�5�   %   $��74��5�   %   $��74��5�   %   �� 5�   &   $
��Wait��60   $ Center=�6<   5�   #   5�      � 5�      $��9��6#   .+�5�   "   ��  �� Orientation$ Vertical=�6-   �� 5�   "   -   5        $��5,0,0,0q�5�      $��App source:��5�      5�      �� 5�   "   $ 
Horizontal=�6-   .+�5�   &   �� -   5        $	 appPath��5�      $	��190��5�      $��
10,22,0,30q�5�      $
��Left=�5�      $

 Wrap=�5�      �� 5�   &   -   5        $ btnOpenFile��5�      $��28��5�      $��28��5      $��3,18,0,0q�5     $	��Top=�5     
������ $��0q�5     .��5  *   �� .+�5  .   �� $��Images/open-file.png��64   �� 5  .   $ Center=�68   5	  +   5
  '   5  #   �� 5  "   -   5        $��5,84,0.2,-13.8q�5     $��Repeat:��5     5     �� 5  "   -   5        $	 
repeatTime��5     $��31��5     $��54,88,0,4.2q�5     $
��Left=�5     $��1��5     5     5     � 5     $��1��6   $��10,10q�6,   .+�5     �� -   5        $ listView��5     $	��593��5     $	��440��5     $��	0,0,0.4,0q�5     ��. $��Disabled=�5"      5$  "   �� 5%  &   .P�5&  *   �� $ 60��5'  !   $��Id��5)  !    5(  !   �� *Id+5)  !   �� 5*  *   $	 260��5+  !   $��Description��5-  !    5,  !   �� *	Content+5-  !   �� 5.  *   $ 90��5/  !   $
��Type��51  !    50  !   �� *Type+51  !   52  '   54     56     � 57     $��2��6   .+�58     �� -   5         ��Name$ countRecord��58  $   $��0 record(s)��69   59     � 5:     $��3��6   .+�5;     �� -   5        $ btnClear��5<     $��3��5=     $��0,0,5,0q�5>     
������ $��0q�5@     .��5A  "   �� .+�5B  &   �� $��Images/clear.png��6,   $ NearestNeighbor=�5C  %   $
��Fill=�5D  %   $��64��5E  %   $��64��5F  %   �� 5G  &   $ Center=�60   Clear6N   5H  #   5I     5J     5K     5M     �� 5N     $��Screen Capture��6   .��5O     �  5P     �� 5Q     $	��120 �6,   �� 5R     $��* �6,   .+�5V     � $��0��6   $��10,10q�6,    5W     �� 5X  "   $	��100 �63   �� 5Y  "   $	��140 �63   �� 5Z  "   $	��100 �63   �� 5[  "   $	��100 �63   �� 5\  "   $	��100 �63   .+�5_     � $��0��6#   .+�5`  "   �� -   5        $ Capture��5a  %   $��0��5b  %   $��0,0,5,0q�5c  %   
������ $��0q�5e  %   .��5g  &   �� .+�5h  *   �� $��Images/capture.png��60   $ NearestNeighbor=�6L   $
��Fill=�5i  %   $��64��5j  %   $��64��5k  %   �� 5l  *   $��0q�64   $ Center=�6?   
Capture!6`   5m  '   5n  #   5p     � 5q     $��1��6#   .+�5r  "   �� -    5        $ Extract_Screenshot_Button��5s  %   $��0��5t  %   $��0,0,5,0q�5u  %   
������ $��0q�5w  %   .��5y  &   �� .+�5z  *   �� $��Images/text.png��60   $ NearestNeighbor=�6I   $
��Fill=�5{  %   $��64��5|  %   $��64��5}  %   �� 5~  *   $��0q�64   $ Center=�6?   Extract Text from picture6q   5  '   5�  #   5�     � 5�     $��2��6#   .+�5�  "   �� -!   5        $ Take_Screenshot��5�  %   $��0��5�  %   $��0,0,5,0q�5�  %   
������ $��0q�5�  %   .��5�  &   �� .+�5�  *   �� $��Images/screenshot.png��60   $ NearestNeighbor=�6O   $
��Fill=�5�  %   $��64��5�  %   $��64��5�  %   �� 5�  *   $��0q�64   $ Center=�6?   Take Screenshot6g   5�  '   5�  #   5�     � 5�     $��3��6#   .+�5�  "   �� -"   5        $ saveScreenshot��5�  %   $��0��5�  %   $��0,0,5,0q�5�  %   
������ $��0q�5�  %   .��5�  &   �� .+�5�  *   �� $��Images/diskette.png��60   $ NearestNeighbor=�6M   $
��Fill=�5�  %   $��64��5�  %   $��64��5�  %   �� 5�  *   $��0q�64   $ Center=�6?   Save Picture6d   5�  '   5�  #   5�     � 5�     $��4��6#   .+�5�  "   �� -#   5        $ Clear��5�  %   $��0��5�  %   $��0,0,5,0q�5�  %   
������ $��0q�5�  %   .��5�  &   �� .+�5�  *   �� $��Images/clear.png��60   $ NearestNeighbor=�6J   $
��Fill=�5�  %   $��64��5�  %   $��64��5�  %   �� 5�  *   $��0q�64   $ Center=�6?   Clear6]   5�  '   5�  #   5�     5�     � 5�     $��1��6   $��10q�6,   .+�5�     �� -$   5         ��Name$ screenShot_picture��5�  $   $
��Left=�6@   $	��512��6[   $	��Top=�6h   $	��650��6�   5�     5�     5�     �� 5�     $��Search word��6   .��5�     � 
������� 5�     �� 5�     $	��550 �6/   .+�5�     � $��0��6   $��0q�6/    5�     �� 5�  "   $��30 �60   �� 5�  "   $��20 �60   �� 5�  "   $��* �60   .+�5�     � $��0��6#   .+�5�  "   �� $��Search��6(   $
��Left=�69   $��Center=�6T   $��25��6o   $
��Auto��6{   $
��Bold"�6�   $��3,5,0,5q�6�   5�     � 5�     -%   5        $��
gridSearch��5�  0   $��2��6#   $
��Auto��6B   $
��Auto��6P    5�  "   �� 5�  &   $��25 �64   �� 5�  &   $��20 �64   �� 5�  &   $	��222 �64   �� 5�  &   $	��350 �64   �� 5�  &   $��15 �64    5�  "   �� 5�  &   $��12 �67   �� 5�  &   $
��150* �67   �� 5�  &   $��12 �67   .+�5�  "   �� $
��Left=�6(   $��1��6C   $��20��6S   $��0��6_   $	��Top=�6l   $��20��6�   $ HighQuality=�6�   $
��None=�6�   $��0,2,0,0q�6�   �� 5�  "   $��
Search for��6(   $��1��6>   $��0��6N   $
��Left=�6[   $��0,0,0,0q�6v   $	��Top=�6�   $	��105��6�   �� 5�  "   -&   5        $��tbSearch��5�  *   $
��Auto��6:   $��1��6G   $��1��6W   � 5�  "   $��2��6'   $��1��64   .+�5�  &   �� -'   5        $��chkWholeWord��5�  /   $��Match whole word��6C   $
��Left=�6^   $��0,10,0,0q�6y   $	��Top=�6�   �� 5�  &   -(   5        $��chkCaseSensitive��5�  /   $��
Match case��6G   $
��Left=�6\   $��0,30,0,0q�6w   $	��Top=�6�   �� 5�  &   $��	Search in��6/   $
��Left=�6B   $��73��6]   $��0,50,0,0q�6i   $	��Top=�6{   $	��198��6�   .��5�  *   � $
��Left=�6/   $��51��6J   $	��Top=�6V   $	��177��6n   .+�5�  .   	� -)   5        $��rbCurrentPage��5�  :   $��current page��6O   $
��Left=�6f   $��0,6,0,0q�6�   $	��Top=�6�   	� 5�  .   -*   5        $��
rbAllPages��5�  :    d� 	IsChecked$
 TrueY�6L   $��	all pages��6]   $
��Left=�6q   $��0,26,0,0q�6�   $	��Top=�6�   5�  +   5�  '   �� 5�  &   -+   5        $��btnFindPrevious��5�  Z   $'��!self:CustomCommand.SearchPrevious��6-   $��Find previous��6q   $
��Left=�6�   $��21��6�   $��	4,128,0,0q�6�   $	��Top=�6�   $��91��6�   �� 5�  &   -,   5        $��btnFindNext��5�  U   $#��self:CustomCommand.SearchNext��6-   $��	Find next��6h   $
��Left=�6|   $��21��6�   $��104,128,0,0q�6�   $	��Top=�6�   $��91��6�   �� 5�  &   --   5        $��	btnSearch��5�  Q   $��self:CustomCommand.Search��6-   $��
Search All��6b   $
��Left=�6w   $��21��6�   $��	4,154,0,0q�6�   $	��Top=�6�   $��76��6�   �� 5�  &   -.   5        $��stSearchProgress��5�  [   $y�Hidden=�60   
�������$��Bottom=�6s   $��0,180,10,10q�6�   .@�5�  *   �� $	��Top=�68   .��5�  .   �� -/   5         ��Name$ lblSearchResults��5�  8   $��5,0,0,0q�6R   $�� ��6c   5�  +   �� 5�  *   $��Stretch=�68   $��Bottom=�6]   .��5   .   � -0   5         �Name$ searchProgressBar��5   :   5  +   5  '   5  #   �� 5  "   -1   5        $��lstSearchResults��5  H   $��1��6+   $��3��6;   $
��Auto��6`   $
��Auto��6m   6{   5     5     5
     5     5     5     �� 5  
   $��0��6   $��Right=�6'   
�������$��Stretch=�6V   $��	1, 0, 1,0q�6r   $��5��6�   � 5  
   $��1��6   $��10,0,0,0q�6    5     �� 5     $��* �6    �� 5     $��30 �6    .+�5     � $��0��6   $��10q�6    .+�5     �� $��1q�6   
���   �._�5       GdPicture14.WPF.GdViewer  -2   5          Name$ 
FileViewer��5  #   $��0q�67      EnableMouseWheel .    ZoomMode$ ZoomModeFitToViewer=�6Z   5     5     � 5     $��	10,0,10,5q�6   $��1��6&    5      �� 5!     $��* �6(   �� 5"     $��* �6'   �� 5#     -3   5         ��Name$ pageCountColumn��5#  '   $��* �6A   �� 5$     $��* �6'   .+�5&     �� -4   5        $��btnPrevious��5&  C   $%��self:CustomCommand.PreviousPage��6   $��0q�6V   $��0��6j   $
��10,0q�6z   
������ .��5'     �� $ 
Horizontal=�6!   $��2,0,0,0q�6:   .+�5(     �� $ HighQuality=�6!   $��Images/left-arrow.png��6O   $
��Fill=�6n   5)     5*     �� 5,     -5   5        $��tbCurrentPage��5,  B   $��Center=�6W    �� TextAlignment$ Center=�6y   $��1��6�   $
��10,0q�6�   �� 5-     -6   5        $��lblPageCount��5-     $
��Left=�6,   $��Center=�6G   $	��/ X��6b   $��26��6q   $
��50,0q�6}   $��2��6�   �� 5/     -7   5        $��btnNext��5/  ?   $!��self:CustomCommand.NextPage��6   $
��10,0q�6N   $��4��6\   $
��Auto��6l   $��0q�6{   
������ .��50     �� $ 
Horizontal=�6!   $��2,0,0,0q�6:   .+�51     �� $ HighQuality=�6!   $��Images/right-arrow.png��6O   $
��Fill=�6o   52     53     54     55     56     59     