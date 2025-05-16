   clear;
   close all;
   clc;
   %%
   PPSI = 600*600; %defines the number of pixels per square inch 	(should match spray card scan resolution)
   SI2SMM = 645.16;  %conversion between square inches to square 	millimeters
   %%
   I = imread('p1.jpg'); % Name of the image
   %%
   [I2, rect] = imcrop(I);
   %%
   figure(1)
   imshow(I2) ;
   %%
   figure(2)
   imagesc(I2);
   imwrite(I2,'Example.jpg')
   %% Percentage Coverage
   figure(3)
   red = I2(:,:,1);
   green = I2(:,:,2);
   blue = I2(:,:,3);
   D = uint8((double(red)+double(green))./2);
   BW = ~im2bw(D,graythresh(D));   %convert to black and white
   PercentCoverage = sum(BW(:))/numel(BW) * 100;
   imagesc(BW)
   axis image
   colormap gray
   title(['Coverage (%) = ',num2str(PercentCoverage)])
   %% Statistics and Saving Data
   figure(4)
   STATS = transpose(cell2mat(struct2cell(regionprops(BW,'Area','Eccentricity')))); %calculate the area and eccentricity (return variable) 
   Avg = mean(STATS(:,1))/PPSI*SI2SMM; %calculate the average size 	in mm^2
   StDev = std(STATS(:,1))/PPSI*SI2SMM;  %calculate the standard 	deviation in mm^2
   hFig = histogram(STATS(:,1)); %plot a histrogram of the droplet 	area in pixels
   edges = hFig.BinEdges;
   values = [hFig.Values];
   ylabel('# of occurrences')
   xlabel('area (pixels)')
   title(sprintf('Mean (mm^2) = %g\n StdDev (mm^2) = %g',Avg,StDev)); 
   folder = 'C:\Users\kshende\Desktop\Rahul Singh'; %File location
   if ~exist(folder, 'dir')
   mkdir(folder);
   end
   baseFileName = 'MyNewFile_data.xlsx'; %File Name
   filename = fullfile(folder, baseFileName);
   P = {num2str(PercentCoverage),Avg,StDev,values};
   sheet1= 1;
   [~,~,ev] = xlsread(filename,sheet1); % ev is cell array 	containing every filled row and column
   empty_row = size(ev,1) + 1; 
   xlRange = ['A',num2str(empty_row)];
   writecell(P,filename,'sheet',sheet1,'Range',xlRange);

