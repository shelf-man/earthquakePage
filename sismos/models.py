from django.db import models


class Sismo(models.Model):
    time = models.CharField(max_length=40, null=True, blank=True)
    year = models.IntegerField(null=True, blank=True, db_index=True)
    latitude = models.FloatField(null=True, blank=True)
    longitude = models.FloatField(null=True, blank=True)
    depth = models.FloatField(null=True, blank=True)
    mag = models.FloatField(null=True, blank=True, db_index=True)
    place = models.CharField(max_length=255, null=True, blank=True)
    mag_category = models.CharField(max_length=50, null=True, blank=True)
    type = models.CharField(max_length=50, null=True, blank=True)

    class Meta:
        db_table = "sismos_sismo"

    def __str__(self):
        return f"{self.time} - Mag {self.mag} - {self.place}"